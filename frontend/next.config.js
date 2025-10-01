/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'export',
  trailingSlash: true,
  distDir: 'out',
  images: {
    domains: ['images.pexels.com'],
    unoptimized: true,
  },
  env: {
    NEXT_PUBLIC_BACKEND_URL: process.env.NEXT_PUBLIC_BACKEND_URL,
  },
  // Remove rewrites for static export
};

module.exports = nextConfig;
