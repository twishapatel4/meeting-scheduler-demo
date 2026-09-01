import Joi from 'joi';
import dotenv from 'dotenv';

dotenv.config();

const envSchema = Joi.object({
  NODE_ENV: Joi.string().valid('development', 'production', 'test').default('development'),
  PORT: Joi.number().default(5000),
  DATABASE_URL: Joi.string().uri().required().messages({
    'string.uri': 'DATABASE_URL must be a valid PostgreSQL URL',
    'any.required': 'DATABASE_URL is required',
  }),
  TOKEN_ENC_KEY: Joi.string()
    .length(64)
    .hex()
    .required()
    .messages({
      'string.length': 'TOKEN_ENC_KEY must be a 32-byte hex string (64 hex chars)',
      'any.required': 'TOKEN_ENC_KEY is required',
    }),
  MS_CLIENT_ID: Joi.string().required(),
  MS_CLIENT_SECRET: Joi.string().required(),
  MS_REDIRECT_URI: Joi.string().uri().default('http://localhost:5000/auth/staff/callback'),
  FRONTEND_URL: Joi.string().uri().default('http://localhost:5173'),
}).unknown(true);

const { error, value } = envSchema.validate(process.env, { abortEarly: false });

if (error) {
  console.error('Invalid environment variables:');
  console.error(error.details.map((d) => `  - ${d.message}`).join('\n'));
  process.exit(1);
}

export const env = value as {
  NODE_ENV: 'development' | 'production' | 'test';
  PORT: number;
  DATABASE_URL: string;
  TOKEN_ENC_KEY: string;
  MS_CLIENT_ID: string;
  MS_CLIENT_SECRET: string;
  MS_REDIRECT_URI: string;
  FRONTEND_URL: string;
};
