import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // 우체국 방문접수 자동화(서버 크롬) — 크롬 바이너리는 묶으면 경로가 깨져 외부 패키지로 둔다
  serverExternalPackages: ["@sparticuz/chromium", "puppeteer-core"],
  // 크롬 압축본(bin/*.br)은 import 가 아니라 파일로 읽혀서 자동 추적이 안 됨 → 해당 API 에 포함
  outputFileTracingIncludes: {
    "/api/repair/epost-pickup": ["./node_modules/@sparticuz/chromium/bin/**"],
  },
};

export default nextConfig;
