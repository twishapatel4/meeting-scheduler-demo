import 'reflect-metadata';
import { plainToInstance } from 'class-transformer';
import { validate } from 'class-validator';
import {
  VisitorRequestDto,
  CreateBookingDto,
  RescheduleDto,
  CancelDto,
  SwapHostDto,
  AvailabilityQueryDto,
} from './booking.dto';

const VALID_UUID = '9c858901-8a57-4791-81fe-4c455b099bc9';
const VALID_UUID_2 = '2d9e4f1a-6b3c-4a2e-9f0d-8b1c2e3f4a5b';

describe('VisitorRequestDto', () => {
  it('passes validation with valid input', async () => {
    const dto = plainToInstance(VisitorRequestDto, {
      visitorEmail: 'visitor@example.com',
      visitorName: 'Jane Visitor',
      requestedStart: '2026-09-02T10:00:00.000Z',
      requestedEnd: '2026-09-02T11:00:00.000Z',
    });
    const errors = await validate(dto);
    expect(errors).toHaveLength(0);
  });

  it('fails when visitorEmail is not a valid email', async () => {
    const dto = plainToInstance(VisitorRequestDto, {
      visitorEmail: 'not-an-email',
      visitorName: 'Jane Visitor',
      requestedStart: '2026-09-02T10:00:00.000Z',
      requestedEnd: '2026-09-02T11:00:00.000Z',
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'visitorEmail')).toBe(true);
  });

  it('fails when requestedStart is not ISO8601', async () => {
    const dto = plainToInstance(VisitorRequestDto, {
      visitorEmail: 'visitor@example.com',
      visitorName: 'Jane Visitor',
      requestedStart: 'not-a-date',
      requestedEnd: '2026-09-02T11:00:00.000Z',
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'requestedStart')).toBe(true);
  });

  it('fails when required fields are missing', async () => {
    const dto = plainToInstance(VisitorRequestDto, {});
    const errors = await validate(dto);
    const properties = errors.map((e) => e.property);
    expect(properties).toEqual(
      expect.arrayContaining(['visitorEmail', 'visitorName', 'requestedStart', 'requestedEnd']),
    );
  });
});

describe('CreateBookingDto', () => {
  it('passes validation with valid input', async () => {
    const dto = plainToInstance(CreateBookingDto, {
      bookingId: VALID_UUID,
      staffId: VALID_UUID_2,
      subject: 'Client Meeting',
    });
    const errors = await validate(dto);
    expect(errors).toHaveLength(0);
  });

  it('fails when bookingId is not a UUID', async () => {
    const dto = plainToInstance(CreateBookingDto, {
      bookingId: 'not-a-uuid',
      staffId: VALID_UUID_2,
      subject: 'Client Meeting',
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'bookingId')).toBe(true);
  });

  it('fails when staffId is not a UUID', async () => {
    const dto = plainToInstance(CreateBookingDto, {
      bookingId: VALID_UUID,
      staffId: '123',
      subject: 'Client Meeting',
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'staffId')).toBe(true);
  });

  it('passes validation when subject is missing (optional, falls back to booking.subject)', async () => {
    const dto = plainToInstance(CreateBookingDto, {
      bookingId: VALID_UUID,
      staffId: VALID_UUID_2,
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'subject')).toBe(false);
  });
});

describe('RescheduleDto', () => {
  it('passes validation with valid input', async () => {
    const dto = plainToInstance(RescheduleDto, {
      newStart: '2026-09-02T10:00:00.000Z',
      newEnd: '2026-09-02T11:00:00.000Z',
    });
    const errors = await validate(dto);
    expect(errors).toHaveLength(0);
  });

  it('fails when newStart is not ISO8601', async () => {
    const dto = plainToInstance(RescheduleDto, {
      newStart: 'tomorrow',
      newEnd: '2026-09-02T11:00:00.000Z',
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'newStart')).toBe(true);
  });
});

describe('CancelDto', () => {
  it('passes validation with valid input', async () => {
    const dto = plainToInstance(CancelDto, {
      comment: 'No longer needed',
      cancelledBy: 'staff-1',
    });
    const errors = await validate(dto);
    expect(errors).toHaveLength(0);
  });

  it('fails when comment is missing', async () => {
    const dto = plainToInstance(CancelDto, { cancelledBy: 'staff-1' });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'comment')).toBe(true);
  });

  it('fails when cancelledBy is not a string', async () => {
    const dto = plainToInstance(CancelDto, { comment: 'reason', cancelledBy: 123 });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'cancelledBy')).toBe(true);
  });
});

describe('SwapHostDto', () => {
  it('passes validation with valid input', async () => {
    const dto = plainToInstance(SwapHostDto, { newStaffId: VALID_UUID });
    const errors = await validate(dto);
    expect(errors).toHaveLength(0);
  });

  it('fails when newStaffId is not a UUID', async () => {
    const dto = plainToInstance(SwapHostDto, { newStaffId: 'nope' });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'newStaffId')).toBe(true);
  });
});

describe('AvailabilityQueryDto', () => {
  it('passes validation without optional staffId', async () => {
    const dto = plainToInstance(AvailabilityQueryDto, {
      start: '2026-09-02T10:00:00.000Z',
      end: '2026-09-02T11:00:00.000Z',
    });
    const errors = await validate(dto);
    expect(errors).toHaveLength(0);
  });

  it('passes validation with a valid optional staffId', async () => {
    const dto = plainToInstance(AvailabilityQueryDto, {
      start: '2026-09-02T10:00:00.000Z',
      end: '2026-09-02T11:00:00.000Z',
      staffId: VALID_UUID,
    });
    const errors = await validate(dto);
    expect(errors).toHaveLength(0);
  });

  it('fails when optional staffId is present but not a UUID', async () => {
    const dto = plainToInstance(AvailabilityQueryDto, {
      start: '2026-09-02T10:00:00.000Z',
      end: '2026-09-02T11:00:00.000Z',
      staffId: 'not-a-uuid',
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'staffId')).toBe(true);
  });

  it('fails when start is not ISO8601', async () => {
    const dto = plainToInstance(AvailabilityQueryDto, {
      start: 'invalid',
      end: '2026-09-02T11:00:00.000Z',
    });
    const errors = await validate(dto);
    expect(errors.some((e) => e.property === 'start')).toBe(true);
  });
});
