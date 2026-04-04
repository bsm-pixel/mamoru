'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save, Plus, X } from 'lucide-react';
import type { TabProps } from '@/app/(dashboard)/settings/page';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

export default function ShippingSettings({ settings, onSave, saving }: TabProps) {
  const [sender, setSender] = useState({ name: '', tel: '', zip: '', addr: '' });
  const [goodsName, setGoodsName] = useState('가위 복원수리');
  const [memoPresets, setMemoPresets] = useState<string[]>([]);
  const [newMemo, setNewMemo] = useState('');
  const [unshippedDays, setUnshippedDays] = useState(2);
  const [reviewDelay, setReviewDelay] = useState(0);
  const [autoPush, setAutoPush] = useState(false);
  const [soundEnabled, setSoundEnabled] = useState(false);
  const [soundTargets, setSoundTargets] = useState({ orders: true, consultations: true, repairs: true });

  useEffect(() => {
    setSender(parse(settings['shipping.sender'], { name: '마모루', tel: '', zip: '', addr: '' }));
    setGoodsName(parse(settings['shipping.default_goods_name'], '가위 복원수리'));
    setMemoPresets(parse(settings['shipping.memo_presets'], []));
    setUnshippedDays(parse(settings['shipping.unshipped_warning_days'], 2));
    setReviewDelay(parse(settings['shipping.review_delay_days'], 0));
    setAutoPush(parse(settings['shipping.auto_push_invoice'], false));
    setSoundEnabled(parse(settings['notifications.sound_enabled'], false));
    setSoundTargets(parse(settings['notifications.sound_targets'], { orders: true, consultations: true, repairs: true }));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'shipping.sender', value: sender },
      { key: 'shipping.default_goods_name', value: goodsName },
      { key: 'shipping.memo_presets', value: memoPresets },
      { key: 'shipping.unshipped_warning_days', value: unshippedDays },
      { key: 'shipping.review_delay_days', value: reviewDelay },
      { key: 'shipping.auto_push_invoice', value: autoPush },
      { key: 'notifications.sound_enabled', value: soundEnabled },
      { key: 'notifications.sound_targets', value: soundTargets },
    ]);
  };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">주문·배송 설정</h2>

      {/* 1. 발송인 정보 */}
      <Field label="발송인 정보" desc="ALPS 송장 생성 시 사용되는 발송인입니다.">
        <div className="grid grid-cols-2 gap-2">
          <input placeholder="이름" value={sender.name} onChange={(e) => setSender({ ...sender, name: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="연락처" value={sender.tel} onChange={(e) => setSender({ ...sender, tel: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="우편번호" value={sender.zip} onChange={(e) => setSender({ ...sender, zip: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="주소" value={sender.addr} onChange={(e) => setSender({ ...sender, addr: e.target.value })}
            className="col-span-2 h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
        </div>
      </Field>

      {/* 2. 기본 상품명 */}
      <Field label="복원수리 송장 상품명" desc="복원수리 출고 시 송장에 찍히는 상품명. 판매 건은 품목명이 자동 입력됩니다.">
        <input value={goodsName} onChange={(e) => setGoodsName(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
          placeholder="[MAMORU] 복원수리" />
      </Field>

      {/* 3. 송장번호 범위 — 별도 UI 필요, 여기선 안내만 */}
      <Field label="송장번호 범위 관리" desc="롯데택배 운송장 번호 구간. 잔여 확인 및 새 구간 등록.">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3">
          송장번호 범위는 <span className="font-semibold">lotte_waybill_config</span> 테이블에서 관리됩니다.
          추후 이 화면에서 직접 관리할 수 있도록 업데이트 예정입니다.
        </p>
      </Field>

      {/* 4. 기본 택배사 — 라벨만 */}
      <Field label="기본 택배사" desc="">
        <p className="text-sm text-neutral-600 bg-neutral-50 rounded-lg px-3 py-2">
          롯데택배 <span className="text-xs text-neutral-400 ml-2">(설정 필요시 기능 구축 필요)</span>
        </p>
      </Field>

      {/* 6. 배송 메모 프리셋 */}
      <Field label="배송 메모 기본 문구" desc="송장 생성 시 드롭다운에서 선택할 수 있는 메모입니다.">
        <div className="space-y-1.5">
          {memoPresets.map((m, i) => (
            <div key={i} className="flex items-center gap-2">
              <span className="flex-1 text-sm bg-neutral-50 rounded-lg px-3 py-1.5">{m}</span>
              <button onClick={() => setMemoPresets(memoPresets.filter((_, j) => j !== i))}
                className="text-neutral-400 hover:text-red-500"><X size={14} /></button>
            </div>
          ))}
          <div className="flex gap-2">
            <input value={newMemo} onChange={(e) => setNewMemo(e.target.value)} placeholder="새 메모 추가"
              className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
              onKeyDown={(e) => { if (e.key === 'Enter' && newMemo.trim()) { setMemoPresets([...memoPresets, newMemo.trim()]); setNewMemo(''); } }} />
            <button onClick={() => { if (newMemo.trim()) { setMemoPresets([...memoPresets, newMemo.trim()]); setNewMemo(''); } }}
              className="px-2 py-1 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
          </div>
        </div>
      </Field>

      {/* 7. 미발송 경고 일수 */}
      <Field label="미발송 경고 일수" desc="결제 후 이 일수 내 미발송 시 주문 목록에서 경고 표시.">
        <div className="flex items-center gap-2">
          <input type="number" value={unshippedDays} onChange={(e) => setUnshippedDays(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={1} />
          <span className="text-sm text-neutral-500">일</span>
        </div>
      </Field>

      {/* 9. 리뷰 요청 대기일 */}
      <Field label="배송완료 후 리뷰 요청 대기일" desc="배송 완료 후 N일 뒤에 리뷰 요청 알림톡 발송. 0이면 즉시.">
        <div className="flex items-center gap-2">
          <input type="number" value={reviewDelay} onChange={(e) => setReviewDelay(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} />
          <span className="text-sm text-neutral-500">일</span>
        </div>
      </Field>

      {/* 10. 배송추적 */}
      <Field label="자동 배송추적" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3">
          매일 자정 Vercel Cron으로 배송 상태를 자동 추적합니다.
          현재 디버깅 진행 중 — 아임웹 배송완료 ↔ TMS 상태 불일치 확인 예정.
        </p>
      </Field>

      {/* 15. 아임웹 송장 자동 연동 */}
      <Field label="아임웹 송장 자동 연동" desc="TMS에서 송장 생성 시 아임웹 주문에도 자동으로 송장번호 입력.">
        <Toggle checked={autoPush} onChange={setAutoPush} />
      </Field>

      {/* 17. 신규 접수 알림음 */}
      <Field label="신규 접수 알림음" desc="새 주문/상담/수리 접수 시 브라우저 알림음을 재생합니다.">
        <Toggle checked={soundEnabled} onChange={setSoundEnabled} />
        {soundEnabled && (
          <div className="mt-2 flex gap-3">
            {(['orders', 'consultations', 'repairs'] as const).map((k) => (
              <label key={k} className="flex items-center gap-1.5 text-xs">
                <input type="checkbox" checked={soundTargets[k]}
                  onChange={(e) => setSoundTargets({ ...soundTargets, [k]: e.target.checked })}
                  className="rounded" />
                {k === 'orders' ? '주문' : k === 'consultations' ? '상담' : '복원수리'}
              </label>
            ))}
          </div>
        )}
      </Field>

      {/* 19. 반품 — Phase 3에서 구현, 여기선 안내 */}
      <Field label="반품 처리 규칙" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3">
          판매 조회 → 반품 처리 기능이 추가됩니다. 반품 시 재고/시리얼이 자동 복귀됩니다.
        </p>
      </Field>

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}>
          <Save size={14} />
          {saving ? '저장 중...' : '저장'}
        </Button>
      </div>
    </div>
  );
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>
      {desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}
      {children}
    </div>
  );
}

function Toggle({ checked, onChange }: { checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <button onClick={() => onChange(!checked)}
      className={`relative w-11 h-6 rounded-full transition ${checked ? 'bg-neutral-900' : 'bg-neutral-200'}`}>
      <span className={`absolute top-0.5 left-0.5 w-5 h-5 rounded-full bg-white shadow transition-transform ${checked ? 'translate-x-5' : ''}`} />
    </button>
  );
}
