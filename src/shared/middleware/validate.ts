import { Request, Response, NextFunction } from 'express';
import { plainToInstance } from 'class-transformer';
import { validate } from 'class-validator';
import { BadRequestError } from '@shared/errors/BadRequestError';

export function validateDto(
  DtoClass: new () => object,
  source: 'body' | 'query' = 'body',
) {
  return async (req: Request, _res: Response, next: NextFunction): Promise<void> => {
    const instance = plainToInstance(DtoClass, req[source]);
    const errors = await validate(instance);
    if (errors.length > 0) {
      const message = errors
        .map((e) => Object.values(e.constraints ?? {}).join(', '))
        .join('; ');
      next(new BadRequestError(message));
      return;
    }
    req[source] = instance as never;
    next();
  };
}
