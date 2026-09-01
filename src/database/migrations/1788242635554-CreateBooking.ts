import { MigrationInterface, QueryRunner } from "typeorm";

export class CreateBooking1788242635554 implements MigrationInterface {
    name = 'CreateBooking1788242635554'

    public async up(queryRunner: QueryRunner): Promise<void> {
        await queryRunner.query(`CREATE TABLE "bookings" ("id" uuid NOT NULL DEFAULT uuid_generate_v4(), "createdAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(), "updatedAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(), "deletedAt" TIMESTAMP WITH TIME ZONE, "visitorEmail" character varying NOT NULL, "visitorName" character varying NOT NULL, "requestedStart" TIMESTAMP WITH TIME ZONE NOT NULL, "requestedEnd" TIMESTAMP WITH TIME ZONE NOT NULL, "status" character varying NOT NULL DEFAULT 'Requested', "msEventId" text, "joinUrl" text, "cancelledAt" TIMESTAMP WITH TIME ZONE, "cancelledBy" text, "staff_id" uuid, CONSTRAINT "PK_bee6805982cc1e248e94ce94957" PRIMARY KEY ("id"))`);
        await queryRunner.query(`ALTER TABLE "bookings" ADD CONSTRAINT "FK_016e3ee9c10ae3cc8ecb3765c4d" FOREIGN KEY ("staff_id") REFERENCES "staff"("id") ON DELETE NO ACTION ON UPDATE NO ACTION`);
    }

    public async down(queryRunner: QueryRunner): Promise<void> {
        await queryRunner.query(`ALTER TABLE "bookings" DROP CONSTRAINT "FK_016e3ee9c10ae3cc8ecb3765c4d"`);
        await queryRunner.query(`DROP TABLE "bookings"`);
    }

}
