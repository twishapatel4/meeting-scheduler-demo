import { AppError } from './AppError';

export class GraphNotFoundError extends AppError {
  constructor(message = 'Graph resource not found (may already be deleted)') {
    super(message, 404, 'GRAPH_NOT_FOUND');
  }
}
