import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // 2026-11-16 화면 점검(npm run check:screen)은 가짜 서버에 붙인 빌드를 따로 만든다 — 평소 빌드(.next)와 섞이지 않게 폴더를 나눈다
  ...(process.env.TMS_CHECK_DIST ? { distDir: process.env.TMS_CHECK_DIST } : {}),
  // 우체국 방문접수 자동화(서버 크롬) — 크롬 바이너리는 묶으면 경로가 깨져 외부 패키지로 둔다
  serverExternalPackages: ["@sparticuz/chromium", "puppeteer-core"],
  // 크롬 압축본(bin/*.br)은 import 가 아니라 파일로 읽혀서 자동 추적이 안 됨 → 해당 API 에 포함
  outputFileTracingIncludes: {
    "/api/repair/epost-pickup": ["./node_modules/@sparticuz/chromium/bin/**"],
  },
  // 2026-11-15 보안 머리글 — 전엔 HSTS 만 있었다
  //   · 다른 사이트가 TMS 를 끼워 넣지 못하게(frame-ancestors / X-Frame-Options). TMS 를 끼워 넣어 쓰는 곳은 없다(고객 페이지는 API 만 부른다)
  //   · 파일 종류 추측 금지(nosniff), 다른 사이트로 나갈 때 주소 전체를 넘기지 않기(Referrer-Policy)
  //   · 검색 엔진에 잡히지 않게(X-Robots-Tag)
  async headers() {
    return [
      {
        source: "/:path*",
        headers: [
          { key: "X-Frame-Options", value: "SAMEORIGIN" },
          { key: "Content-Security-Policy", value: "frame-ancestors 'self'" },
          { key: "X-Content-Type-Options", value: "nosniff" },
          { key: "Referrer-Policy", value: "strict-origin-when-cross-origin" },
          { key: "X-Robots-Tag", value: "noindex, nofollow" },
        ],
      },
    ];
  },
};

export default nextConfig;
