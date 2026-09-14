import { Entity, Column, ManyToOne, JoinColumn } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Booking } from './booking.entity';

export type InsightStatus = 'pending' | 'processing' | 'completed' | 'not_found' | 'failed';

export interface ActionItem {
  id?: string;
  title?: string;
  text: string;
  owner?: string | null;
  dueDate?: string | null;
}

export interface MeetingNote {
  title?: string;
  text?: string;
  subpoints?: string[];
}

@Entity('meeting_ai_insights')
export class MeetingAiInsight extends BaseEntity {
  @ManyToOne(() => Booking, { onDelete: 'CASCADE', nullable: false })
  @JoinColumn({ name: 'booking_id' })
  booking: Booking;

  @Column({ name: 'booking_id' })
  bookingId: string;

  @Column({ type: 'text', nullable: true })
  graphInsightId: string | null;

  @Column({ type: 'text', default: 'microsoft_graph' })
  provider: string;

  @Column({ type: 'text', nullable: true })
  summary: string | null;

  @Column({ type: 'jsonb', nullable: true })
  notes: MeetingNote[] | null;

  @Column({ type: 'jsonb', nullable: true })
  actionItems: ActionItem[] | null;

  @Column({ type: 'varchar', length: 50, default: 'pending' })
  status: InsightStatus;

  @Column({ type: 'text', nullable: true })
  errorMessage: string | null;

  @Column({ type: 'timestamptz', nullable: true })
  generatedAt: Date | null;
}
