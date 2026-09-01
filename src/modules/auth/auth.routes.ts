import { Router } from 'express';
import { Container } from 'typedi';
import { AuthController } from './auth.controller';
import { asyncHandler } from '@shared/utils/asyncHandler';

const router = Router();
const controller = Container.get(AuthController);

router.get('/staff/login', asyncHandler(controller.login));
router.get('/staff/callback', asyncHandler(controller.callback));

export { router as authRouter };
