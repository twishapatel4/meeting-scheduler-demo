import { IsString, IsEmail, IsISO8601, IsUUID, IsOptional } from 'class-validator';

export class VisitorRequestDto {
  @IsEmail()
  visitorEmail: string;

  @IsString()
  visitorName: string;

  @IsISO8601()
  requestedStart: string;

  @IsISO8601()
  requestedEnd: string;

  @IsOptional()
  @IsString()
  subject?: string;
}

export class CreateBookingDto {
  @IsUUID()
  bookingId: string;

  @IsUUID()
  staffId: string;

  @IsOptional()
  @IsString()
  subject?: string;
}

export class RescheduleDto {
  @IsISO8601()
  newStart: string;

  @IsISO8601()
  newEnd: string;

  @IsOptional()
  @IsString()
  subject?: string;
}

export class CancelDto {
  @IsString()
  comment: string;

  @IsString()
  cancelledBy: string;
}

export class SwapHostDto {
  @IsUUID()
  newStaffId: string;
}

export class AvailabilityQueryDto {
  @IsISO8601()
  start: string;

  @IsISO8601()
  end: string;

  @IsOptional()
  @IsUUID()
  staffId?: string;
}
