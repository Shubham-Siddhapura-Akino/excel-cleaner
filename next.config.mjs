/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  serverExternalPackages: ['exceljs', 'jszip', 'fastify', '@fastify/cors', '@fastify/multipart'],
};

export default nextConfig;
