import { buildFastify } from './app.js';

const PORT = parseInt(process.env.FASTIFY_PORT || process.env.PORT || '3001', 10);

async function start() {
  const app = await buildFastify({
    logger: {
      transport: {
        target: 'pino-pretty',
        options: {
          translateTime: 'HH:MM:ss Z',
          ignore: 'pid,hostname',
        },
      },
    },
  }).catch(async () => {
    // Fallback without pino-pretty if not installed
    return await buildFastify({ logger: true });
  });

  try {
    const address = await app.listen({ port: PORT, host: '0.0.0.0' });
    console.log('\n======================================================');
    console.log(`  🚀 Fastify Backend is running!`);
    console.log(`  📡 API Address: ${address}`);
    console.log(`  🔗 Health Check: ${address}/api/health`);
    console.log('======================================================\n');
  } catch (err) {
    app.log.error(err);
    process.exit(1);
  }
}

start();
