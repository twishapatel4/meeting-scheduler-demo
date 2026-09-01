import { AppError } from './AppError';

export class GraphAuthError extends AppError {
  constructor(message = 'Microsoft Graph rejected the access token') {
    super(message, 401, 'GRAPH_AUTH_ERROR');
  }
}
