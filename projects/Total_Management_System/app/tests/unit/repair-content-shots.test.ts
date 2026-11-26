import { describe, it, expect } from 'vitest';
import { pendingShots } from '@/components/repairs/content-candidate';

/**
 * 복원수리 콘텐츠 후보 — 촬영 확인 줄을 언제 띄우나 (2026-11-26, 작업실 19번)
 * 후보이고 아직 안 찍었을 때만. 작업 전 = 작업에 손대기 전~작업 중, 작업 후 = 작업 중~작업완료(출고 전)
 */
const base = { id: 'r1', content_candidate: true, content_symptom: null, content_before_shot: false, content_after_shot: false, content_consent: null } as const;

describe('촬영 확인 줄', () => {
  it('후보가 아니면 아무것도 안 띄운다', () => {
    expect(pendingShots({ ...base, content_candidate: false, status: 'repairing' })).toEqual([]);
  });
  it('입고 전후에는 작업 전 촬영만', () => {
    expect(pendingShots({ ...base, status: 'cost_notified' })).toEqual(['before']);
  });
  it('작업 중에는 작업 전 · 작업 후 둘 다(작업 전을 놓친 경우까지)', () => {
    expect(pendingShots({ ...base, status: 'repairing' })).toEqual(['before', 'after']);
  });
  it('찍은 것은 다시 안 띄운다', () => {
    expect(pendingShots({ ...base, status: 'repairing', content_before_shot: true })).toEqual(['after']);
    expect(pendingShots({ ...base, status: 'ready_to_ship', content_before_shot: true, content_after_shot: true })).toEqual([]);
  });
  it('출고 뒤 · 취소 건에는 안 띄운다', () => {
    expect(pendingShots({ ...base, status: 'shipped' })).toEqual([]);
    expect(pendingShots({ ...base, status: 'cancelled' })).toEqual([]);
  });
});
