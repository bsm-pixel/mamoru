import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { scanCallSites, makeNeeds, knownTemplates, ALWAYS_SENT, type MakeReads } from './scan';

/**
 * 알림톡 종류별 점검 — 템플릿 칸이 읽는 값을 TMS 가 실제로 보내는가 (2026-11-22)
 *
 * 계기: 출장 접수 알림톡에 희망 요일 · 시간대가 빈칸으로 고객에게 나갔다(작업실 t-261121-field-wish-alimtalk).
 *   Make 분기는 템플릿 칸을 어떤 이름에서 읽는데, TMS 가 다른 이름으로 보내면 그 칸은 빈칸이 된다.
 *
 * 비교하는 두 표
 *   - Make 쪽: tests/alimtalk/make-reads.json — `node tests/alimtalk/snapshot-make.mjs` 로 Make 에서 새로 뽑는다
 *   - TMS 쪽: src 안의 sendNotification(...) 호출을 글자로 읽는다(tests/alimtalk/scan.ts). 실제로 보내지 않는다.
 *
 * 실패하면: 「어느 신호 · 어느 파일 · 어느 칸」이 빠졌는지 나온다. 고치거나, 일부러 둔 것이면 KNOWN_GAPS 에 이유와 함께 적는다.
 */
const SRC = path.resolve(__dirname, '../../src');
const reads: MakeReads = JSON.parse(fs.readFileSync(path.join(__dirname, 'make-reads.json'), 'utf8'));
const needs = makeNeeds(reads);
const KNOWN = knownTemplates(SRC);

/** 신호를 코드만으로 못 따라가는 곳 — 실제로 넘어오는 신호를 사람이 확인해 적어 둔다 */
const DYNAMIC: Record<string, string[]> = {
  // 상담 상세 [확정 알림 다시 보내기] (app/(dashboard)/consultations/[id]/page.tsx)
  'app/api/consultation/notify/route.ts': ['confirmed'],
  // 복원수리 [비용 안내 보내기] (components/repairs/sidebar-action-card.tsx)
  'app/api/repair/[id]/notify/route.ts': ['as_cost_notice'],
  // EVENT 입금 안내 · 입금 확인 (재고판매 LS 는 10-22 폐기 — Make 분기 없음)
  'app/api/events/[id]/route.ts': ['event_payment_notice', 'event_payment_confirmed'],
  // 아임웹 취소 · 반품 · 교환 웹훅 (lib/imweb/claim-notify.ts 의 ACTIONS 표)
  'lib/imweb/claim-notify.ts': ['imweb_cancel_requested', 'imweb_cancel_completed', 'imweb_return_requested', 'imweb_return_approved', 'imweb_return_completed', 'imweb_exchange_requested'],
};

/** `...extraData` 처럼 화면이 덧붙여 보내는 칸 — 실제로 넘기는 칸을 적어 두면 그곳도 점검한다 */
const CALLER_EXTRA: Record<string, string[]> = {
  // 상담 상세 [알림톡 재발송] 은 extraData 없이 부른다
  'app/api/consultation/notify/route.ts': [],
  // 복원수리 [비용 안내] 는 금액 3칸을 덧붙인다
  'app/api/repair/[id]/notify/route.ts': ['as_amount', 'shipping_amount', 'total_amount'],
};

/**
 * 알고 있는 빈칸 — 2026-11-22 첫 점검에서 나온 것. 대표님이 번호로 고르면 그것만 고치고 여기서 지운다.
 * 형식: '신호 · 파일 · 칸'
 */
const KNOWN_GAPS = new Set<string>([
  // 1) 복원수리 출고 알림톡 — Make 는 #{tracking_number} 를 tracking_number 에서 읽는데 TMS 는 tracking 으로 보낸다
  'as_shipped · app/api/cron/track-delivery/route.ts · tracking_number',
  'as_shipped · app/api/repair/[id]/merged-ship/route.ts · tracking_number',
  'as_shipped · app/api/repair/[id]/route.ts · tracking_number',
  'as_shipped · lib/notification/combined-shipped.ts · tracking_number',
  // 2) 매장방문 확정 알림톡 [재발송] — 일정 변경 버튼 주소(change_request_link)를 안 보낸다
  'confirmed · app/api/consultation/notify/route.ts · change_request_link',
]);

const sites = scanCallSites(SRC).map((s) => {
  const dyn = DYNAMIC[s.file];
  const templates = dyn ? dyn : s.templates.filter((t) => KNOWN.has(t));
  const extra = CALLER_EXTRA[s.file];
  return extra ? { ...s, templates, keys: [...s.keys, ...extra], open: false } : { ...s, templates };
});

/** 이번 점검으로 찾은 빈칸 전부 */
function findGaps() {
  const gaps: { key: string; signal: string; file: string; line: number; field: string; partial: boolean; open: boolean }[] = [];
  for (const s of sites) {
    for (const t of s.templates) {
      const n = needs.get(t);
      if (!n) continue;
      const have = new Set([...ALWAYS_SENT, ...s.keys]);
      for (const f of n.some) {
        if (have.has(f)) continue;
        if (s.open) continue; // ...extraData 등으로 넘어올 수 있다 — 아래 '알 수 없음' 목록으로 따로 보인다
        gaps.push({ key: `${t} · ${s.file} · ${f}`, signal: t, file: s.file, line: s.line, field: f, partial: !n.all.has(f), open: s.open });
      }
    }
  }
  return gaps;
}

describe('알림톡 — 템플릿 칸이 읽는 값을 TMS 가 보낸다', () => {
  it('보내는 곳마다 신호를 알아낼 수 있다 (모르면 DYNAMIC 에 적는다)', () => {
    const unknown = sites.filter((s) => s.templates.length === 0).map((s) => `${s.file}:${s.line}`);
    expect(unknown, '신호를 못 따라간 곳 — DYNAMIC 에 실제 신호를 적어 주세요').toEqual([]);
  });

  it('새 빈칸이 없다 (알고 있는 빈칸은 KNOWN_GAPS)', () => {
    const fresh = findGaps().filter((g) => !KNOWN_GAPS.has(g.key))
      .map((g) => `${g.signal} @ ${g.file}:${g.line} — 「${g.field}」 를 안 보냄${g.partial ? ' (일부 분기만 읽음)' : ''}`);
    expect(fresh, '템플릿 칸이 읽는데 TMS 가 보내지 않는 값').toEqual([]);
  });

  it('KNOWN_GAPS 에 이미 고쳐진 항목이 남아 있지 않다', () => {
    const now = new Set(findGaps().map((g) => g.key));
    const stale = [...KNOWN_GAPS].filter((k) => !now.has(k));
    expect(stale, '고쳐졌으면 KNOWN_GAPS 에서 지워 주세요').toEqual([]);
  });

  it('값이 빈 글자(\'\')로 고정된 칸을 템플릿이 읽지 않는다', () => {
    const bad: string[] = [];
    for (const s of sites) for (const t of s.templates) {
      const n = needs.get(t);
      for (const f of s.emptyLiterals) if (n?.some.has(f)) bad.push(`${t} @ ${s.file}:${s.line} — 「${f}」 = ''`);
    }
    expect(bad).toEqual([]);
  });
});

/** 점검 결과 표 — 실패 여부와 무관하게 보고용으로 만든다(npm run check 출력에는 안 나옴) */
export function report() {
  const sentSignals = new Set(sites.flatMap((s) => s.templates));
  return {
    gaps: findGaps(),
    openSites: sites.filter((s) => s.open).map((s) => ({ at: `${s.file}:${s.line}`, templates: s.templates })),
    makeOnly: [...needs.keys()].filter((k) => !sentSignals.has(k)),
    codeOnly: [...sentSignals].filter((k) => !needs.has(k)),
  };
}
