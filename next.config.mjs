/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  output: 'export',
  trailingSlash: true,
  assetPrefix: '.',
  images: {
    unoptimized: true,
  },
};

export default nextConfig;
