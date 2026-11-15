'use client';

// 2026-11-15 가장 바깥에서 오류가 났을 때의 마지막 안전망 — 하얀 화면 대신 다시 시도 버튼
export default function GlobalError({ reset }: { error: Error & { digest?: string }; reset: () => void }) {
  return (
    <html lang="ko">
      <body style={{ margin: 0, fontFamily: 'Pretendard, system-ui, sans-serif', background: '#FAF9F7', color: '#1A1A1A' }}>
        <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24 }}>
          <div style={{ textAlign: 'center' }}>
            <p style={{ fontSize: 14, fontWeight: 700 }}>문제가 생겼습니다</p>
            <p style={{ fontSize: 12, color: '#8A8580', marginTop: 6 }}>잠시 뒤 다시 시도해 주세요.</p>
            <button type="button" onClick={reset} style={{ marginTop: 20, height: 40, padding: '0 20px', borderRadius: 8, border: 0, background: '#1A1A1A', color: '#fff', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}>다시 시도</button>
          </div>
        </div>
      </body>
    </html>
  );
}
