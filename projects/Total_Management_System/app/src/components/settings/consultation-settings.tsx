'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save, Plus, X } from 'lucide-react';
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

interface ClosedDate { id?: string; date: string; reason: string; }

export default function ConsultationSettings({ settings, onSave, saving }: TabProps) {
  const supabase = createClient();
  const [cs, setCs] = useState<ConsultSettings>({
    start_hour: 10, end_hour: 20, duration_min: 60, step_min: 10,
    disabled_weekdays: [0], field_buffer_before: 90, field_buffer_after: 90,
  });
  const [closedDates, setClosedDates] = useState<ClosedDate[]>([]);
  const [newDate, setNewDate] = useState('');
  const [newReason, setNewReason] = useState('');
  const [reminder24h, setReminder24h] = useState(true);
  const [reminder2h, setReminder2h] = useState(true);
  const [autoReview, setAutoReview] = useState(false);
  const [gmailNotify, setGmailNotify] = useState(true);
  const [durationByType, setDurationByType] = useState({ store_visit: 60, field_request: 90, talk_consult: 0 });
  const [changeDeadline, setChangeDeadline] = useState(0);
  const [storeAddress, setStoreAddress] = useState('');

  useEffect(() => {
    // consultation_settings 테이블에서 직접 조회
    (async () => {
      const { data } = await (supabase as any).from('consultation_settings').select('*').eq('id', 'default').single();
      if (data) {
        setCs({
          start_hour: data.start_hour ?? 10, end_hour: data.end_hour ?? 20,
          duration_min: data.duration_min ?? 60, step_min: data.step_min ?? 10,
          disabled_weekdays: data.disabled_weekdays ?? [0],
          field_buffer_before: data.field_buffer_before ?? 90, field_buffer_after: data.field_buffer_after ?? 90,
        });
      }
      const { data: cd } = await (supabase as any).from('closed_dates').select('*').order('date', { ascending: true });
      setClosedDates((cd || []) as ClosedDate[]);
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
    // consultation_settings 업데이트 (직접 DB)
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

  const addClosedDate = async () => {
    if (!newDate) return;
    const { error } = await (supabase as any).from('closed_dates').insert({ date: newDate, reason: newReason || '휴무' });
    if (error) { toast.error('휴무일 추가 실패'); return; }
    setClosedDates([...closedDates, { date: newDate, reason: newReason || '휴무' }]);
    setNewDate(''); setNewReason('');
    toast.success('휴무일 추가 완료');
  };

  const removeClosedDate = async (cd: ClosedDate) => {
    if (cd.id) await (supabase as any).from('closed_dates').delete().eq('id', cd.id);
    setClosedDates(closedDates.filter((d) => d.date !== cd.date));
  };

  const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">상담 관리 설정</h2>

      {/* 1-2. 영업시간 */}
      <Field label="영업 시간" desc="고객 접수 폼의 예약 가능 시간대에 반영됩니다.">
        <div className="flex items-center gap-2">
          <input type="number" value={cs.start_hour} onChange={(e) => setCs({ ...cs, start_hour: Number(e.target.value) })}
            className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" min={0} max={23} />
          <span className="text-sm">시 ~</span>
          <input type="number" value={cs.end_hour} onChange={(e) => setCs({ ...cs, end_hour: Number(e.target.value) })}
            className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" min={0} max={23} />
          <span className="text-sm">시</span>
        </div>
      </Field>

      {/* 2. 휴무 요일 */}
      <Field label="휴무 요일" desc="체크한 요일은 접수 폼에서 예약 불가.">
        <div className="flex gap-2">
          {WEEKDAYS.map((d, i) => (
            <button key={i}
              onClick={() => {
                const next = cs.disabled_weekdays.includes(i)
                  ? cs.disabled_weekdays.filter((w) => w !== i)
                  : [...cs.disabled_weekdays, i];
                setCs({ ...cs, disabled_weekdays: next });
              }}
              className={`w-9 h-9 rounded-lg text-sm font-medium transition ${
                cs.disabled_weekdays.includes(i) ? 'bg-red-500 text-white' : 'bg-neutral-100 text-neutral-600'
              }`}
            >{d}</button>
          ))}
        </div>
      </Field>

      {/* 3. 특별 휴무일 */}
      <Field label="특별 휴무일" desc="연휴/개인 사정 등 임시 휴무.">
        <div className="space-y-1.5 mb-2">
          {closedDates.map((cd, i) => (
            <div key={i} className="flex items-center gap-2 text-sm">
              <span className="font-mono">{cd.date}</span>
              <span className="text-neutral-500">{cd.reason}</span>
              <button onClick={() => removeClosedDate(cd)} className="text-neutral-400 hover:text-red-500 ml-auto"><X size={14} /></button>
            </div>
          ))}
        </div>
        <div className="flex gap-2">
          <input type="date" value={newDate} onChange={(e) => setNewDate(e.target.value)}
            className="h-8 px-2 rounded-lg border border-neutral-200 text-sm" />
          <input value={newReason} onChange={(e) => setNewReason(e.target.value)} placeholder="사유"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm" />
          <button onClick={addClosedDate} className="px-2 py-1 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
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
