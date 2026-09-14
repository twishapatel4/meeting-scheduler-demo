import { Router } from 'express';
import { Container } from 'typedi';
import { BookingController } from './booking.controller';
import { asyncHandler } from '@shared/utils/asyncHandler';
import { validateDto } from '@shared/middleware/validate';
import { VisitorRequestDto, CreateBookingDto, RescheduleDto, CancelDto, SwapHostDto } from './booking.dto';

const router = Router();
const controller = Container.get(BookingController);

router.get('/', asyncHandler(controller.list));
router.post('/visitor-request', validateDto(VisitorRequestDto), asyncHandler(controller.visitorRequest));
router.post('/', validateDto(CreateBookingDto), asyncHandler(controller.create));
router.patch('/:id', validateDto(RescheduleDto), asyncHandler(controller.reschedule));
router.post('/:id/cancel', validateDto(CancelDto), asyncHandler(controller.cancel));
router.delete('/:id', asyncHandler(controller.deleteHard));
router.post('/:id/swap', validateDto(SwapHostDto), asyncHandler(controller.swap));
router.get('/:id/insights', asyncHandler(controller.getInsights));
router.post('/:id/fetch-insights', asyncHandler(controller.fetchInsights));

export { router as bookingRouter };
