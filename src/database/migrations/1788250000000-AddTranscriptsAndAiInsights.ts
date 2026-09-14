import { MigrationInterface, QueryRunner } from 'typeorm';

export class AddTranscriptsAndAiInsights1788250000000 implements MigrationInterface {
  name = 'AddTranscriptsAndAiInsights1788250000000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`ALTER TABLE "bookings" ADD COLUMN "onlineMeetingId" text`);

    await queryRunner.query(`
      CREATE TABLE "meeting_transcripts" (
        "id" uuid NOT NULL DEFAULT uuid_generate_v4(),
        "createdAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(),
        "updatedAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(),
        "deletedAt" TIMESTAMP WITH TIME ZONE,
        "booking_id" uuid NOT NULL,
        "graphTranscriptId" text,
        "onlineMeetingId" text,
        "contentType" text NOT NULL DEFAULT 'text/vtt',
        "rawContent" text,
        "plainText" text,
        "segments" jsonb,
        "status" character varying(50) NOT NULL DEFAULT 'pending',
        "errorMessage" text,
        "availableAt" TIMESTAMP WITH TIME ZONE,
        CONSTRAINT "PK_meeting_transcripts_id" PRIMARY KEY ("id"),
        CONSTRAINT "FK_meeting_transcripts_booking" FOREIGN KEY ("booking_id") REFERENCES "bookings"("id") ON DELETE CASCADE
      )
    `);

    await queryRunner.query(`
      CREATE TABLE "meeting_ai_insights" (
        "id" uuid NOT NULL DEFAULT uuid_generate_v4(),
        "createdAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(),
        "updatedAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(),
        "deletedAt" TIMESTAMP WITH TIME ZONE,
        "booking_id" uuid NOT NULL,
        "graphInsightId" text,
        "provider" text NOT NULL DEFAULT 'microsoft_graph',
        "summary" text,
        "notes" jsonb,
        "actionItems" jsonb,
        "status" character varying(50) NOT NULL DEFAULT 'pending',
        "errorMessage" text,
        "generatedAt" TIMESTAMP WITH TIME ZONE,
        CONSTRAINT "PK_meeting_ai_insights_id" PRIMARY KEY ("id"),
        CONSTRAINT "FK_meeting_ai_insights_booking" FOREIGN KEY ("booking_id") REFERENCES "bookings"("id") ON DELETE CASCADE
      )
    `);
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE "meeting_ai_insights"`);
    await queryRunner.query(`DROP TABLE "meeting_transcripts"`);
    await queryRunner.query(`ALTER TABLE "bookings" DROP COLUMN "onlineMeetingId"`);
  }
}
