import { Router } from 'express';
import { Container } from 'typedi';
import { StaffController } from './staff.controller';
import { asyncHandler } from '@shared/utils/asyncHandler';
import { validateDto } from '@shared/middleware/validate';
import { AvailabilityQueryDto } from '@modules/booking/booking.dto';

const router = Router();
const controller = Container.get(StaffController);

router.get('/', asyncHandler(controller.list));
router.get('/:id/availability', validateDto(AvailabilityQueryDto, 'query'), asyncHandler(controller.availability));

export { router as staffRouter };
