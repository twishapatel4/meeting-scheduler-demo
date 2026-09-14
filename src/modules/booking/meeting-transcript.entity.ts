import { Entity, Column, ManyToOne, JoinColumn } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Booking } from './booking.entity';

export type TranscriptStatus = 'pending' | 'processing' | 'available' | 'not_found' | 'failed';

export interface TranscriptSegment {
  startTime: string;
  endTime: string;
  speaker: string;
  text: string;
}

@Entity('meeting_transcripts')
export class MeetingTranscript extends BaseEntity {
  @ManyToOne(() => Booking, { onDelete: 'CASCADE', nullable: false })
  @JoinColumn({ name: 'booking_id' })
  booking: Booking;

  @Column({ name: 'booking_id' })
  bookingId: string;

  @Column({ type: 'text', nullable: true })
  graphTranscriptId: string | null;

  @Column({ type: 'text', nullable: true })
  onlineMeetingId: string | null;

  @Column({ type: 'text', default: 'text/vtt' })
  contentType: string;

  @Column({ type: 'text', nullable: true })
  rawContent: string | null;

  @Column({ type: 'text', nullable: true })
  plainText: string | null;

  @Column({ type: 'jsonb', nullable: true })
  segments: TranscriptSegment[] | null;

  @Column({ type: 'varchar', length: 50, default: 'pending' })
  status: TranscriptStatus;

  @Column({ type: 'text', nullable: true })
  errorMessage: string | null;

  @Column({ type: 'timestamptz', nullable: true })
  availableAt: Date | null;
}
