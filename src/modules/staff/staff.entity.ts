import { Entity, Column, ManyToOne, JoinColumn } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Organization } from '@modules/organization/organization.entity';

@Entity('staff')
export class Staff extends BaseEntity {
  @ManyToOne(() => Organization, { nullable: false })
  @JoinColumn({ name: 'organization_id' })
  organization: Organization;

  @Column({ unique: true })
  email: string;

  @Column({ type: 'text', nullable: true })
  refreshTokenEnc: string | null;

  @Column({ type: 'text', nullable: true })
  accessToken: string | null;

  @Column({ type: 'timestamptz', nullable: true })
  tokenExpiresAt: Date | null;

  @Column({ default: 'enterprise' })
  accountType: string; // 'enterprise' | 'small_business' | 'personal'

  @Column({ default: false })
  connected: boolean;
}
