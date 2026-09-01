import 'reflect-metadata';
import { IsEmail, IsString } from 'class-validator';
import { Request, Response, NextFunction } from 'express';
import { validateDto } from './validate';
import { BadRequestError } from '@shared/errors/BadRequestError';

class SampleDto {
  @IsEmail()
  email: string;

  @IsString()
  name: string;
}

function buildReqRes(body: unknown): { req: Request; res: Response } {
  const req = { body, query: {} } as unknown as Request;
  const res = {} as Response;
  return { req, res };
}

describe('validateDto middleware', () => {
  it('calls next() with no error when the body is valid', async () => {
    const { req, res } = buildReqRes({ email: 'a@example.com', name: 'Alice' });
    const next: NextFunction = jest.fn();

    await validateDto(SampleDto)(req, res, next);

    expect(next).toHaveBeenCalledWith();
    expect(req.body).toBeInstanceOf(SampleDto);
  });

  it('calls next(err) with a BadRequestError when the body is invalid', async () => {
    const { req, res } = buildReqRes({ email: 'not-an-email' });
    const next: NextFunction = jest.fn();

    await validateDto(SampleDto)(req, res, next);

    expect(next).toHaveBeenCalledTimes(1);
    const err = (next as jest.Mock).mock.calls[0][0];
    expect(err).toBeInstanceOf(BadRequestError);
    expect(err.message).toEqual(expect.stringContaining('email'));
  });

  it('validates req.query when source is "query"', async () => {
    const req = { body: {}, query: { email: 'not-an-email', name: 'Bob' } } as unknown as Request;
    const res = {} as Response;
    const next: NextFunction = jest.fn();

    await validateDto(SampleDto, 'query')(req, res, next);

    expect(next).toHaveBeenCalledTimes(1);
    const err = (next as jest.Mock).mock.calls[0][0];
    expect(err).toBeInstanceOf(BadRequestError);
  });
});
