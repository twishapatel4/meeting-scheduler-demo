import { Service } from 'typedi';
import { Request, Response } from 'express';
import { StaffService } from './staff.service';
import { BookingService } from '@modules/booking/booking.service';
import { sendSuccess } from '@shared/utils/apiResponse';
import { NotFoundError } from '@shared/errors/NotFoundError';
import { sanitizeStaff } from './staff.sanitize';

@Service()
export class StaffController {
  constructor(
    private readonly staffService: StaffService,
    private readonly bookingService: BookingService,
  ) {}

  list = async (_req: Request, res: Response): Promise<void> => {
    const staff = await this.staffService.listAll();
    sendSuccess(res, staff.map(sanitizeStaff));
  };

  availability = async (req: Request, res: Response): Promise<void> => {
    const staff = await this.staffService.findById(req.params.id);
    if (!staff) throw new NotFoundError('Staff not found');

    const start = new Date(req.query.start as string);
    const end = new Date(req.query.end as string);
    const result = await this.bookingService.checkAvailability(staff.id, start, end);
    sendSuccess(res, result);
  };
}
