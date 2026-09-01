import { Entity, Column, OneToMany } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Staff } from '@modules/staff/staff.entity';

@Entity('organizations')
export class Organization extends BaseEntity {
  @Column()
  name: string;

  @Column({ unique: true })
  domain: string;

  @OneToMany(() => Staff, (staff) => staff.organization)
  staff: Staff[];
}
