export class HttpError extends Error {
  readonly statusCode: number;
  readonly code: string;

  constructor(statusCode: number, code: string, message: string) {
    super(message);
    this.statusCode = statusCode;
    this.code = code;
    this.name = 'HttpError';
  }
}

export const unauthorized = (message = 'Unauthorized') =>
  new HttpError(401, 'unauthorized', message);

export const forbidden = (message = 'Forbidden') =>
  new HttpError(403, 'forbidden', message);

export const notFound = (message = 'Not found') =>
  new HttpError(404, 'not_found', message);

export const badRequest = (message = 'Bad request') =>
  new HttpError(400, 'bad_request', message);

export const renderFailed = (message = 'Render failed') =>
  new HttpError(502, 'render_failed', message);
