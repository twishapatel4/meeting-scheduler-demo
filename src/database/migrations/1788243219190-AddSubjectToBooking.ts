import { MigrationInterface, QueryRunner } from "typeorm";

export class AddSubjectToBooking1788243219190 implements MigrationInterface {
    name = 'AddSubjectToBooking1788243219190'

    public async up(queryRunner: QueryRunner): Promise<void> {
        await queryRunner.query(`ALTER TABLE "bookings" ADD "subject" character varying NOT NULL DEFAULT 'Meeting'`);
    }

    public async down(queryRunner: QueryRunner): Promise<void> {
        await queryRunner.query(`ALTER TABLE "bookings" DROP COLUMN "subject"`);
    }

}
