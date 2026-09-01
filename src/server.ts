import 'reflect-metadata';
import { app } from './app';
import { AppDataSource } from '@config/database';
import { env } from '@config/env';
import { logger } from '@shared/utils/logger';

const startServer = async (): Promise<void> => {
  await AppDataSource.initialize();
  logger.info('Database connected');

  const server = app.listen(env.PORT, () => {
    logger.info(`Server running on port ${env.PORT}`, { env: env.NODE_ENV });
  });

  const shutdown = (signal: string): void => {
    logger.info(`${signal} received — shutting down`);
    server.close(async () => {
      await AppDataSource.destroy();
      process.exit(0);
    });
  };

  process.on('SIGTERM', () => shutdown('SIGTERM'));
  process.on('SIGINT', () => shutdown('SIGINT'));
};

startServer().catch((err: Error) => {
  logger.error('Failed to start server', { message: err.message, stack: err.stack });
  process.exit(1);
});
