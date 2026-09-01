import { Service } from 'typedi';
import { Repository } from 'typeorm';
import { AppDataSource } from '@config/database';
import { createGraphClient } from '@shared/graph-client';
import { GraphNotFoundError } from '@shared/errors/GraphNotFoundError';
import { NotFoundError } from '@shared/errors/NotFoundError';
import { BadRequestError } from '@shared/errors/BadRequestError';
import { logger } from '@shared/utils/logger';
import { Booking } from './booking.entity';
import { StaffService } from '@modules/staff/staff.service';
import { TokenService } from '@modules/auth/token.service';

interface VisitorRequestInput {
  visitorEmail: string;
  visitorName: string;
  requestedStart: Date;
  requestedEnd: Date;
}

@Service()
export class BookingService {
  constructor(
    private readonly staffService: StaffService,
    private readonly tokenService: TokenService,
  ) {}

  private get repo(): Repository<Booking> {
    return AppDataSource.getRepository(Booking);
  }

  listAll(): Promise<Booking[]> {
    return this.repo.find({ relations: ['staff', 'staff.organization'], order: { createdAt: 'DESC' } });
  }

  private async findOrThrow(id: string): Promise<Booking> {
    const booking = await this.repo.findOne({
      where: { id },
      relations: ['staff', 'staff.organization'],
    });
    if (!booking) throw new NotFoundError('Booking not found');
    return booking;
  }

  createVisitorRequest(data: VisitorRequestInput): Promise<Booking> {
    return this.repo.save(this.repo.create({ ...data, status: 'Requested' }));
  }

  // spec §4 Step 1
  async checkAvailability(
    staffId: string,
    start: Date,
    end: Date,
  ): Promise<{ availabilityView: string; scheduleItems: unknown[] }> {
    const staff = await this.staffService.findById(staffId);
    if (!staff) throw new NotFoundError('Staff not found');

    const accessToken = await this.tokenService.getValidAccessToken(staffId);
    const graph = createGraphClient(accessToken);

    const response = await graph.post('/me/calendar/getSchedule', {
      schedules: [staff.email],
      startTime: { dateTime: start.toISOString().replace('Z', ''), timeZone: 'UTC' },
      endTime: { dateTime: end.toISOString().replace('Z', ''), timeZone: 'UTC' },
      availabilityViewInterval: 30,
    });

    const result = response.data.value[0];
    return { availabilityView: result.availabilityView, scheduleItems: result.scheduleItems };
  }

  // spec §4 Step 2
  async createBooking(bookingId: string, staffId: string, subject: string): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    const staff = await this.staffService.findById(staffId);
    if (!staff) throw new NotFoundError('Staff not found');

    const accessToken = await this.tokenService.getValidAccessToken(staffId);
    const graph = createGraphClient(accessToken);

    const response = await graph.post('/me/events', {
      subject,
      body: { contentType: 'HTML', content: 'Discussion about organization inquiry.' },
      start: { dateTime: booking.requestedStart.toISOString().replace('Z', ''), timeZone: 'UTC' },
      end: { dateTime: booking.requestedEnd.toISOString().replace('Z', ''), timeZone: 'UTC' },
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      attendees: [
        { emailAddress: { address: booking.visitorEmail, name: booking.visitorName }, type: 'required' },
      ],
    });

    booking.staff = staff;
    booking.subject = subject;
    booking.msEventId = response.data.id;
    booking.joinUrl = response.data.onlineMeeting?.joinUrl ?? null;
    booking.status = 'Scheduled';
    return this.repo.save(booking);
  }

  // spec §4 Step 3
  async reschedule(bookingId: string, newStart: Date, newEnd: Date, subject?: string): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    if (!booking.staff || !booking.msEventId) {
      throw new BadRequestError('Booking has no scheduled event to reschedule');
    }

    const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
    const graph = createGraphClient(accessToken);

    try {
      await graph.patch(`/me/events/${booking.msEventId}`, {
        start: { dateTime: newStart.toISOString().replace('Z', ''), timeZone: 'UTC' },
        end: { dateTime: newEnd.toISOString().replace('Z', ''), timeZone: 'UTC' },
        ...(subject !== undefined ? { subject } : {}),
      });
    } catch (err) {
      if (err instanceof GraphNotFoundError) {
        logger.warn('Event already deleted in Outlook — cannot reschedule', { bookingId });
      }
      throw err;
    }

    booking.requestedStart = newStart;
    booking.requestedEnd = newEnd;
    if (subject !== undefined) {
      booking.subject = subject;
    }
    booking.status = 'Rescheduled';
    return this.repo.save(booking);
  }

  // spec §4 Step 4 (cancel — notifies visitor)
  async cancel(bookingId: string, comment: string, cancelledBy: string): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    if (!booking.staff || !booking.msEventId) {
      throw new BadRequestError('Booking has no scheduled event to cancel');
    }

    const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
    const graph = createGraphClient(accessToken);

    try {
      await graph.post(`/me/events/${booking.msEventId}/cancel`, { comment });
    } catch (err) {
      if (!(err instanceof GraphNotFoundError)) throw err;
      logger.warn('Event already deleted in Outlook — treating as cancelled', { bookingId });
    }

    booking.status = 'Cancelled';
    booking.cancelledAt = new Date();
    booking.cancelledBy = cancelledBy;
    return this.repo.save(booking);
  }

  // spec §4 Step 4 (delete — no notification)
  async deleteHard(bookingId: string): Promise<void> {
    const booking = await this.findOrThrow(bookingId);
    if (booking.staff && booking.msEventId) {
      const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
      const graph = createGraphClient(accessToken);
      try {
        await graph.delete(`/me/events/${booking.msEventId}`);
      } catch (err) {
        if (!(err instanceof GraphNotFoundError)) throw err;
      }
    }
    await this.repo.remove(booking);
  }

  assertSameOrganization(orgIdA: string, orgIdB: string): void {
    if (orgIdA !== orgIdB) {
      throw new BadRequestError('Swap-host is only allowed between staff in the same organization');
    }
  }

  // spec §4 Step 5 — cancel with Person A's token, rebook with Person B's token
  async swapHost(bookingId: string, newStaffId: string): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    if (!booking.staff || !booking.msEventId) {
      throw new BadRequestError('Booking has no scheduled event to swap');
    }

    const newStaff = await this.staffService.findById(newStaffId);
    if (!newStaff) throw new NotFoundError('New staff not found');

    this.assertSameOrganization(booking.staff.organization.id, newStaff.organization.id);

    const oldAccessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
    const oldGraph = createGraphClient(oldAccessToken);
    try {
      await oldGraph.post(`/me/events/${booking.msEventId}/cancel`, {
        comment: 'Apologies, this meeting has been reassigned to a different host.',
      });
    } catch (err) {
      if (!(err instanceof GraphNotFoundError)) throw err;
      logger.warn('Old event already deleted in Outlook — proceeding with rebook', { bookingId });
    }

    const newAccessToken = await this.tokenService.getValidAccessToken(newStaff.id);
    const newGraph = createGraphClient(newAccessToken);
    const response = await newGraph.post('/me/events', {
      subject: booking.subject,
      body: { contentType: 'HTML', content: 'Discussion about organization inquiry.' },
      start: { dateTime: booking.requestedStart.toISOString().replace('Z', ''), timeZone: 'UTC' },
      end: { dateTime: booking.requestedEnd.toISOString().replace('Z', ''), timeZone: 'UTC' },
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      attendees: [
        { emailAddress: { address: booking.visitorEmail, name: booking.visitorName }, type: 'required' },
      ],
    });

    booking.staff = newStaff;
    booking.msEventId = response.data.id;
    booking.joinUrl = response.data.onlineMeeting?.joinUrl ?? null;
    booking.status = 'Swapped';
    return this.repo.save(booking);
  }
}
