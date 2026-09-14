import { Entity, Column, ManyToOne, JoinColumn } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Staff } from '@modules/staff/staff.entity';

export type BookingStatus = 'Requested' | 'Scheduled' | 'Rescheduled' | 'Swapped' | 'Cancelled';

@Entity('bookings')
export class Booking extends BaseEntity {
  @Column()
  visitorEmail: string;

  @Column()
  visitorName: string;

  @Column({ default: 'Meeting' })
  subject: string;

  @Column({ type: 'timestamptz' })
  requestedStart: Date;

  @Column({ type: 'timestamptz' })
  requestedEnd: Date;

  @Column({ default: 'Requested' })
  status: BookingStatus;

  @ManyToOne(() => Staff, { nullable: true })
  @JoinColumn({ name: 'staff_id' })
  staff: Staff | null;

  @Column({ type: 'text', nullable: true })
  msEventId: string | null;

  @Column({ type: 'text', nullable: true })
  joinUrl: string | null;

  @Column({ type: 'text', nullable: true })
  onlineMeetingId: string | null;

  @Column({ type: 'timestamptz', nullable: true })
  cancelledAt: Date | null;

  @Column({ type: 'text', nullable: true })
  cancelledBy: string | null;
}

