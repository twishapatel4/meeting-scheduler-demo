import { Service } from 'typedi';
import { Request, Response } from 'express';
import { BookingService } from './booking.service';
import { Booking } from './booking.entity';
import { sendSuccess } from '@shared/utils/apiResponse';
import { sanitizeStaff } from '@modules/staff/staff.sanitize';
import {
  VisitorRequestDto,
  CreateBookingDto,
  RescheduleDto,
  CancelDto,
  SwapHostDto,
} from './booking.dto';

// Every Booking response carries the full Staff relation (needed by the
// admin dashboard's "swap to <email>" UI) — strip its encrypted/raw tokens
// before this ever leaves the server.
function sanitizeBooking(booking: Booking) {
  return { ...booking, staff: booking.staff ? sanitizeStaff(booking.staff) : null };
}

@Service()
export class BookingController {
  constructor(private readonly bookingService: BookingService) {}

  list = async (_req: Request, res: Response): Promise<void> => {
    const bookings = await this.bookingService.listAll();
    sendSuccess(res, bookings.map(sanitizeBooking));
  };

  visitorRequest = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as VisitorRequestDto;
    const booking = await this.bookingService.createVisitorRequest({
      visitorEmail: dto.visitorEmail,
      visitorName: dto.visitorName,
      requestedStart: new Date(dto.requestedStart),
      requestedEnd: new Date(dto.requestedEnd),
    });
    sendSuccess(res, sanitizeBooking(booking), 'Booking request received', 201);
  };

  create = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as CreateBookingDto;
    const booking = await this.bookingService.createBooking(dto.bookingId, dto.staffId, dto.subject);
    sendSuccess(res, sanitizeBooking(booking), 'Meeting scheduled');
  };

  reschedule = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as RescheduleDto;
    const booking = await this.bookingService.reschedule(
      req.params.id,
      new Date(dto.newStart),
      new Date(dto.newEnd),
      dto.subject,
    );
    sendSuccess(res, sanitizeBooking(booking), 'Booking rescheduled');
  };

  cancel = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as CancelDto;
    const booking = await this.bookingService.cancel(req.params.id, dto.comment, dto.cancelledBy);
    sendSuccess(res, sanitizeBooking(booking), 'Booking cancelled');
  };

  deleteHard = async (req: Request, res: Response): Promise<void> => {
    await this.bookingService.deleteHard(req.params.id);
    sendSuccess(res, {}, 'Booking deleted', 200);
  };

  swap = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as SwapHostDto;
    const booking = await this.bookingService.swapHost(req.params.id, dto.newStaffId);
    sendSuccess(res, sanitizeBooking(booking), 'Host swapped');
  };
}
