import { defineConfig } from 'vitest/config';
import path from 'node:path';

// 자동 점검(2026-11-15) — 운영 자료·외부 서비스에 닿지 않는다. 실행: npm run check
export default defineConfig({
  resolve: { alias: { '@': path.resolve(__dirname, 'src') } },
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts'],
    setupFiles: ['tests/helpers/setup.ts'],
    testTimeout: 20000,
  },
});
