import { Request, Response, NextFunction } from 'express';
import { QueryFailedError } from 'typeorm';
import { AppError } from '@shared/errors/AppError';
import { logger } from '@shared/utils/logger';
import { sendError } from '@shared/utils/apiResponse';

export const errorHandler = (
  err: Error,
  req: Request,
  res: Response,
  _next: NextFunction,
): void => {
  if (err instanceof AppError) {
    sendError(res, err.message, err.statusCode, err.errorCode);
    return;
  }

  if (err instanceof QueryFailedError) {
    const dbErr = err as QueryFailedError & { code?: string };
    if (dbErr.code === '23505') {
      sendError(res, 'A record with the given data already exists', 409, 'CONFLICT');
      return;
    }
  }

  const axiosErr = err as Error & { response?: { data?: unknown; status?: number } };
  logger.error('Unhandled error', {
    path: req.path,
    message: err.message,
    stack: err.stack,
    graphResponse: axiosErr.response?.data,
  });
  sendError(res, 'An unexpected error occurred', 500, 'INTERNAL_SERVER_ERROR');
};
