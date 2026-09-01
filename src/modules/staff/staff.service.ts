import { Service } from 'typedi';
import { Repository } from 'typeorm';
import { AppDataSource } from '@config/database';
import { encrypt, decrypt } from '@shared/crypto';
import { Staff } from './staff.entity';

interface TokenUpdate {
  refreshToken: string;
  accessToken: string;
  expiresAt: Date;
}

@Service()
export class StaffService {
  private get repo(): Repository<Staff> {
    return AppDataSource.getRepository(Staff);
  }

  findById(id: string): Promise<Staff | null> {
    return this.repo.findOne({ where: { id }, relations: ['organization'] });
  }

  findByEmail(email: string): Promise<Staff | null> {
    return this.repo.findOne({ where: { email }, relations: ['organization'] });
  }

  listAll(): Promise<Staff[]> {
    return this.repo.find({ relations: ['organization'], order: { createdAt: 'DESC' } });
  }

  async upsertByEmail(
    email: string,
    organizationId: string,
    accountType: string,
  ): Promise<Staff> {
    const existing = await this.repo.findOne({ where: { email } });
    if (existing) return existing;
    return this.repo.save(
      this.repo.create({
        email,
        organization: { id: organizationId } as Staff['organization'],
        accountType,
        connected: false,
      }),
    );
  }

  async updateTokens(id: string, tokens: TokenUpdate): Promise<void> {
    await this.repo.update(
      { id },
      {
        refreshTokenEnc: encrypt(tokens.refreshToken),
        accessToken: tokens.accessToken,
        tokenExpiresAt: tokens.expiresAt,
        connected: true,
      },
    );
  }

  async markDisconnected(id: string): Promise<void> {
    await this.repo.update({ id }, { connected: false });
  }

  getDecryptedRefreshToken(staff: Staff): string | null {
    if (!staff.refreshTokenEnc) return null;
    return decrypt(staff.refreshTokenEnc);
  }
}
