import { Service } from 'typedi';
import { Repository } from 'typeorm';
import { AppDataSource } from '@config/database';
import { Organization } from './organization.entity';

@Service()
export class OrganizationService {
  private get repo(): Repository<Organization> {
    return AppDataSource.getRepository(Organization);
  }

  create(data: Partial<Organization>): Promise<Organization> {
    return this.repo.save(this.repo.create(data));
  }

  findById(id: string): Promise<Organization | null> {
    return this.repo.findOne({ where: { id } });
  }

  findByDomain(domain: string): Promise<Organization | null> {
    return this.repo.findOne({ where: { domain } });
  }

  async findOrCreateByDomain(domain: string, name: string): Promise<Organization> {
    const existing = await this.findByDomain(domain);
    if (existing) return existing;
    return this.create({ domain, name });
  }
}
