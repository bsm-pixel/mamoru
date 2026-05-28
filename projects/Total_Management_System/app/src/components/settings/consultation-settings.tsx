'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';
import toast from 'react-hot-toast';
import type { TabProps } from '@/app/(dashboard)/settings/page';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

interface ConsultSettings {
  start_hour: number; end_hour: number; duration_min: number; step_min: number;
  disabled_weekdays: number[]; field_buffer_before: number; field_buffer_after: number;
}

export default function ConsultationSettings({ settings, onSave, saving }: TabProps) {
  const supabase = createClient();
  const [cs, setCs] = useState<ConsultSettings>({
    start_hour: 10, end_hour: 20, duration_min: 60, step_min: 10,
    disabled_weekdays: [0], field_buffer_before: 90, field_buffer_after: 90,
  });
  const [reminder24h, setReminder24h] = useState(true);
  const [reminder2h, setReminder2h] = useState(true);
  const [autoReview, setAutoReview] = useState(false);
  const [gmailNotify, setGmailNotify] = useState(true);
  const [durationByType, setDurationByType] = useState({ store_visit: 60, field_request: 90, talk_consult: 0 });
  const [changeDeadline, setChangeDeadline] = useState(0);
  const [storeAddress, setStoreAddress] = useState('');

  useEffect(() => {
    // consultation_settings 테이블에서 직접 조회 (disabled_weekdays는 그대로 보존 — 달력 관리에서 변경)
    (async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data } = await (supabase as any).from('consultation_settings').select('*').eq('id', 'default').single();
      if (data) {
        setCs({
          start_hour: data.start_hour ?? 10, end_hour: data.end_hour ?? 20,
          duration_min: data.duration_min ?? 60, step_min: data.step_min ?? 10,
          disabled_weekdays: data.disabled_weekdays ?? [0],
          field_buffer_before: data.field_buffer_before ?? 90, field_buffer_after: data.field_buffer_after ?? 90,
        });
      }
    })();
    setReminder24h(parse(settings['consultation.reminder_24h_enabled'], true));
    setReminder2h(parse(settings['consultation.reminder_2h_enabled'], true));
    setAutoReview(parse(settings['consultation.auto_review_request'], false));
    setGmailNotify(parse(settings['consultation.gmail_notify'], true));
    setDurationByType(parse(settings['consultation.duration_by_type'], { store_visit: 60, field_request: 90, talk_consult: 0 }));
    setChangeDeadline(parse(settings['consultation.change_deadline_hours'], 0));
    setStoreAddress(parse(settings['business.store_address'], ''));
  }, [settings]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSave = async () => {
    // consultation_settings 업데이트 (직접 DB) — disabled_weekdays는 cs 객체 유지값 그대로 보존
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { error } = await (supabase as any)
      .from('consultation_settings')
      .upsert({ id: 'default', ...cs, updated_at: new Date().toISOString() });
    if (error) { toast.error('상담 설정 저장 실패: ' + error.message); return; }

    // system_settings 업데이트
    onSave([
      { key: 'consultation.reminder_24h_enabled', value: reminder24h },
      { key: 'consultation.reminder_2h_enabled', value: reminder2h },
      { key: 'consultation.auto_review_request', value: autoReview },
      { key: 'consultation.gmail_notify', value: gmailNotify },
      { key: 'consultation.duration_by_type', value: durationByType },
      { key: 'consultation.change_deadline_hours', value: changeDeadline },
      { key: 'business.store_address', value: storeAddress },
    ]);
  };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">상담 관리 설정</h2>

      {/* 영업시간 + 휴무 + 시간대 차단 → 달력 관리 화면으로 통합 이전 (078 / 096) */}
      <Field label="영업시간 · 휴무 관리" desc="영업 기본시간, 정기 휴무 요일, 특정 날짜 휴무, 30분 단위 시간대 차단을 달력 관리 한 화면에서 통합 관리합니다.">
        <a
          href="/consultations/calendar"
          className="inline-flex items-center gap-2 px-3 h-9 rounded-lg bg-stone-900 text-white text-sm font-medium hover:bg-stone-800 transition"
        >
          달력 관리로 이동
        </a>
      </Field>

      {/* 4. 출장 버퍼 */}
      <Field label="출장 전후 버퍼 시간" desc="출장 상담 전후로 차단되는 시간.">
        <div className="flex items-center gap-2">
          <span className="text-sm">전</span>
          <input type="number" value={cs.field_buffer_before} onChange={(e) => setCs({ ...cs, field_buffer_before: Number(e.target.value) })}
            className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" step={10} />
          <span className="text-sm">분 / 후</span>
          <input type="number" value={cs.field_buffer_after} onChange={(e) => setCs({ ...cs, field_buffer_after: Number(e.target.value) })}
            className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" step={10} />
          <span className="text-sm">분</span>
        </div>
      </Field>

      {/* 5. 상담 시간 단위 */}
      <Field label="상담 시간 단위" desc="한 건의 기본 상담 시간.">
        <select value={cs.duration_min} onChange={(e) => setCs({ ...cs, duration_min: Number(e.target.value) })}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          {[30, 45, 60, 90, 120].map((m) => <option key={m} value={m}>{m}분</option>)}
        </select>
      </Field>

      {/* 6. 슬롯 간격 */}
      <Field label="슬롯 간격" desc="예약 가능 시간의 간격.">
        <select value={cs.step_min} onChange={(e) => setCs({ ...cs, step_min: Number(e.target.value) })}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          {[10, 15, 20, 30, 60].map((m) => <option key={m} value={m}>{m}분</option>)}
        </select>
      </Field>

      {/* 7. 리마인더 on/off */}
      <Field label="리마인더 발송" desc="알림톡 리마인더를 발송할지 선택합니다.">
        <div className="flex gap-4">
          <label className="flex items-center gap-2 text-sm">
            <input type="checkbox" checked={reminder24h} onChange={(e) => setReminder24h(e.target.checked)} className="rounded" />
            24시간 전
          </label>
          <label className="flex items-center gap-2 text-sm">
            <input type="checkbox" checked={reminder2h} onChange={(e) => setReminder2h(e.target.checked)} className="rounded" />
            2시간 전
          </label>
        </div>
      </Field>

      {/* 10. 리뷰 요청 자동 */}
      <Field label="상담 완료 후 리뷰 요청 자동 발송" desc="">
        <Toggle checked={autoReview} onChange={setAutoReview} />
      </Field>

      {/* 11. 매장 주소 */}
      <Field label="매장 주소 (지도 중심점)" desc="상담 관리 지도의 중심 좌표 및 출장 거리 기준.">
        <input value={storeAddress} onChange={(e) => setStoreAddress(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" placeholder="서울 강남구..." />
      </Field>

      {/* 14. Gmail 알림 */}
      <Field label="상담 접수 시 Gmail 알림" desc="">
        <Toggle checked={gmailNotify} onChange={setGmailNotify} />
      </Field>

      {/* 16. 유형별 기본 시간 */}
      <Field label="상담 유형별 기본 시간" desc="">
        <div className="space-y-2">
          {([['store_visit', '매장방문'], ['field_request', '출장'], ['talk_consult', '온라인상담']] as const).map(([k, label]) => (
            <div key={k} className="flex items-center gap-2">
              <span className="text-sm w-16">{label}</span>
              <input type="number" value={durationByType[k]} onChange={(e) => setDurationByType({ ...durationByType, [k]: Number(e.target.value) })}
                className="w-16 h-8 px-2 rounded-lg border border-neutral-200 text-sm text-center" />
              <span className="text-xs text-neutral-500">분 (0=제한없음)</span>
            </div>
          ))}
        </div>
      </Field>

      {/* 20. 예약 변경 제한 */}
      <Field label="예약 변경 가능 시간" desc="상담 N시간 전까지만 변경/취소 가능. 0이면 제한 없음.">
        <div className="flex items-center gap-2">
          <input type="number" value={changeDeadline} onChange={(e) => setChangeDeadline(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} />
          <span className="text-sm text-neutral-500">시간 전</span>
        </div>
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
