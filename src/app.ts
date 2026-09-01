import 'reflect-metadata';
import express from 'express';
import helmet from 'helmet';
import cors from 'cors';
import { env } from '@config/env';
import { errorHandler } from '@shared/middleware/errorHandler';
import { healthRouter } from './health/health.routes';
import { authRouter } from '@modules/auth/auth.routes';
import { bookingRouter } from '@modules/booking/booking.routes';
import { staffRouter } from '@modules/staff/staff.routes';

const app = express();

app.use(helmet());
app.use(cors({ origin: env.FRONTEND_URL, credentials: true }));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use('/health', healthRouter);
app.use('/auth', authRouter);
app.use('/api/v1/bookings', bookingRouter);
app.use('/api/v1/staff', staffRouter);

app.use(errorHandler);

export { app };
