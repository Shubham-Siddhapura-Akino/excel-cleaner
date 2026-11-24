import { buildFastify } from '../../server/app.js';

let cachedApp = null;

async function getFastify() {
  if (!cachedApp) {
    cachedApp = await buildFastify();
    await cachedApp.ready();
  }
  return cachedApp;
}

export const config = {
  api: {
    bodyParser: false,
    responseLimit: false,
    externalResolver: true,
  },
};

export default async function handler(req, res) {
  const app = await getFastify();
  app.server.emit('request', req, res);
}
