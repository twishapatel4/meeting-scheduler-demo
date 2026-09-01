import { Service } from 'typedi';
import { Request, Response } from 'express';
import { TokenService } from './token.service';
import { StaffService } from '@modules/staff/staff.service';
import { OrganizationService } from '@modules/organization/organization.service';
import { env } from '@config/env';
import { logger } from '@shared/utils/logger';

@Service()
export class AuthController {
  constructor(
    private readonly tokenService: TokenService,
    private readonly staffService: StaffService,
    private readonly organizationService: OrganizationService,
  ) {}

  login = async (_req: Request, res: Response): Promise<void> => {
    const url = await this.tokenService.getAuthCodeUrl();
    res.redirect(url);
  };

  callback = async (req: Request, res: Response): Promise<void> => {
    const code = req.query.code as string | undefined;
    if (!code) {
      res.redirect(`${env.FRONTEND_URL}/admin?connect=error`);
      return;
    }

    const { accessToken, refreshToken, expiresOn, email } =
      await this.tokenService.exchangeCodeForTokens(code);

    const domain = email.split('@')[1] ?? 'unknown.local';
    const accountType = domain === 'outlook.com' || domain === 'hotmail.com' ? 'personal' : 'enterprise';
    const organization = await this.organizationService.findOrCreateByDomain(domain, domain);

    const staff = await this.staffService.upsertByEmail(email, organization.id, accountType);
    await this.staffService.updateTokens(staff.id, { accessToken, refreshToken, expiresAt: expiresOn });

    logger.info('Staff connected', { email });
    res.redirect(`${env.FRONTEND_URL}/admin?connect=success`);
  };
}
