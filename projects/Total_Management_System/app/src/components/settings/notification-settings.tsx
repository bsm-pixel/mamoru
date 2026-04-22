'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save } from 'lucide-react';
import type { TabProps } from '@/app/(dashboard)/settings/page';
import GoogleCalendarSettings from '@/components/settings/google-calendar-settings';
import BannerSettings from '@/components/settings/banner-settings';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

export default function NotificationSettings({ settings, onSave, saving }: TabProps) {
  const [masterEnabled, setMasterEnabled] = useState(true);
  const [consultationReceived, setConsultationReceived] = useState(true);
  const [repairReceived, setRepairReceived] = useState(true);
  const [repairCostNotice, setRepairCostNotice] = useState(true);
  const [repairPayment, setRepairPayment] = useState(true);
  const [repairShipped, setRepairShipped] = useState(true);
  const [reviewRequest, setReviewRequest] = useState(true);
  const [webhookConsultation, setWebhookConsultation] = useState('');
  const [webhookAsReceived, setWebhookAsReceived] = useState('');
  const [webhookRepair, setWebhookRepair] = useState('');
  // 내 푸시 알림 (사장님이 받는 것)
  const [pushConsultation, setPushConsultation] = useState(true);
  const [pushFieldRequest, setPushFieldRequest] = useState(true);
  const [pushTalkReceived, setPushTalkReceived] = useState(true);
  const [pushFieldConfirmed, setPushFieldConfirmed] = useState(true);
  const [pushFieldReschedule, setPushFieldReschedule] = useState(true);
  const [pushRepairReceived, setPushRepairReceived] = useState(true);
  const [pushReviewSubmitted, setPushReviewSubmitted] = useState(true);
  const [pushOrderReceived, setPushOrderReceived] = useState(true);

  useEffect(() => {
    setMasterEnabled(parse(settings['notifications.master_enabled'], true));
    setConsultationReceived(parse(settings['notifications.consultation_received'], true));
    setRepairReceived(parse(settings['notifications.repair_received'], true));
    setRepairCostNotice(parse(settings['notifications.repair_cost_notice'], true));
    setRepairPayment(parse(settings['notifications.repair_payment_confirmed'], true));
    setRepairShipped(parse(settings['notifications.repair_shipped'], true));
    setReviewRequest(parse(settings['notifications.review_request'], true));
    setWebhookConsultation(parse(settings['notifications.webhook_consultation'], ''));
    setWebhookAsReceived(parse(settings['notifications.webhook_as_received'], ''));
    setWebhookRepair(parse(settings['notifications.webhook_repair'], ''));
    setPushConsultation(parse(settings['push.consultation_received'], true));
    setPushFieldRequest(parse(settings['push.field_request'], true));
    setPushTalkReceived(parse(settings['push.talk_received'], true));
    setPushFieldConfirmed(parse(settings['push.field_confirmed'], true));
    setPushFieldReschedule(parse(settings['push.field_reschedule'], true));
    setPushRepairReceived(parse(settings['push.repair_received'], true));
    setPushReviewSubmitted(parse(settings['push.review_submitted'], true));
    setPushOrderReceived(parse(settings['push.order_received'], true));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'notifications.master_enabled', value: masterEnabled },
      { key: 'notifications.consultation_received', value: consultationReceived },
      { key: 'notifications.repair_received', value: repairReceived },
      { key: 'notifications.repair_cost_notice', value: repairCostNotice },
      { key: 'notifications.repair_payment_confirmed', value: repairPayment },
      { key: 'notifications.repair_shipped', value: repairShipped },
      { key: 'notifications.review_request', value: reviewRequest },
      { key: 'notifications.webhook_consultation', value: webhookConsultation },
      { key: 'notifications.webhook_as_received', value: webhookAsReceived },
      { key: 'notifications.webhook_repair', value: webhookRepair },
      { key: 'push.consultation_received', value: pushConsultation },
      { key: 'push.field_request', value: pushFieldRequest },
      { key: 'push.talk_received', value: pushTalkReceived },
      { key: 'push.field_confirmed', value: pushFieldConfirmed },
      { key: 'push.field_reschedule', value: pushFieldReschedule },
      { key: 'push.repair_received', value: pushRepairReceived },
      { key: 'push.review_submitted', value: pushReviewSubmitted },
      { key: 'push.order_received', value: pushOrderReceived },
    ]);
  };

  const PUSH_ITEMS = [
    { key: 'push_consultation', label: '상담 접수 (매장방문)', state: pushConsultation, setter: setPushConsultation },
    { key: 'push_field', label: '상담 접수 (출장)', state: pushFieldRequest, setter: setPushFieldRequest },
    { key: 'push_talk', label: '상담 접수 (톡상담)', state: pushTalkReceived, setter: setPushTalkReceived },
    { key: 'push_field_confirmed', label: '출장 일정 확정 ✅ (고객 선택)', state: pushFieldConfirmed, setter: setPushFieldConfirmed },
    { key: 'push_field_resched', label: '출장 일정 재요청 🔄 (고객 재요청)', state: pushFieldReschedule, setter: setPushFieldReschedule },
    { key: 'push_repair', label: '복원수리 접수', state: pushRepairReceived, setter: setPushRepairReceived },
    { key: 'push_review', label: '고객 리뷰 작성 ⭐', state: pushReviewSubmitted, setter: setPushReviewSubmitted },
    { key: 'push_order', label: '아임웹 주문 접수 📦', state: pushOrderReceived, setter: setPushOrderReceived },
  ];

  const NOTIF_ITEMS = [
    { key: 'consultation_received', label: '상담 접수 확인', state: consultationReceived, setter: setConsultationReceived },
    { key: 'repair_received', label: '복원수리 접수 확인', state: repairReceived, setter: setRepairReceived },
    { key: 'repair_cost_notice', label: '비용안내', state: repairCostNotice, setter: setRepairCostNotice },
    { key: 'repair_payment', label: '입금확인', state: repairPayment, setter: setRepairPayment },
    { key: 'repair_shipped', label: '출고', state: repairShipped, setter: setRepairShipped },
    { key: 'review_request', label: '리뷰 요청', state: reviewRequest, setter: setReviewRequest },
  ];

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">알림·연동 설정</h2>

      {/* ── Google Calendar 연동 ── */}
      <GoogleCalendarSettings />

      {/* ── 아임웹 배너/팝업 관리 ── */}
      <BannerSettings />

      {/* ── 내 푸시 알림 (사장님이 받는 것) ── */}
      <div className="rounded-lg border border-neutral-200 bg-warm-ivory p-4 space-y-3">
        <div>
          <h3 className="text-sm font-bold text-indigo-black">📱 내 푸시 알림</h3>
          <p className="text-xs text-neutral-500 mt-0.5">고객 행동이 발생하면 사장님 디바이스로 푸시 알림을 보냅니다. (크롬/모바일 앱)</p>
        </div>
        <div className="space-y-2">
          {PUSH_ITEMS.map((item) => (
            <div key={item.key} className="flex items-center justify-between py-1">
              <span className="text-sm">{item.label}</span>
              <Toggle checked={item.state} onChange={item.setter} />
            </div>
          ))}
        </div>
      </div>

      <div className="pt-4 border-t border-neutral-100">
        <h3 className="text-sm font-bold text-neutral-700 mb-3">💬 고객 알림톡 발송</h3>
      </div>

      {/* 1. 마스터 on/off */}
      <Field label="알림톡 전체 on/off" desc="끄면 모든 알림톡 발송이 즉시 중단됩니다. 점검/테스트 시 사용.">
        <div className="flex items-center gap-3">
          <Toggle checked={masterEnabled} onChange={setMasterEnabled} />
          <span className={`text-sm font-medium ${masterEnabled ? 'text-green-600' : 'text-red-500'}`}>
            {masterEnabled ? '활성' : '비활성 — 모든 알림 중단'}
          </span>
        </div>
      </Field>

      {/* 2~7. 개별 on/off */}
      <Field label="알림 유형별 on/off" desc="개별 알림을 세밀하게 제어합니다.">
        <div className={`space-y-3 ${!masterEnabled ? 'opacity-40 pointer-events-none' : ''}`}>
          {NOTIF_ITEMS.map((item) => (
            <div key={item.key} className="flex items-center justify-between">
              <span className="text-sm">{item.label}</span>
              <Toggle checked={item.state} onChange={item.setter} />
            </div>
          ))}
        </div>
      </Field>

      {/* Make 웹훅 URL — 3개 시나리오 */}
      <Field label="Make 웹훅 URL (상담)" desc="상담 접수/확정/취소/리마인더/리뷰 등">
        <input value={webhookConsultation} onChange={(e) => setWebhookConsultation(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      <Field label="Make 웹훅 URL (AS접수)" desc="복원수리 접수 안내 (방문수거/직접발송/카운터)">
        <input value={webhookAsReceived} onChange={(e) => setWebhookAsReceived(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      <Field label="Make 웹훅 URL (AS상태변경)" desc="입고확인/입금안내/출고&송장/취소/만족도">
        <input value={webhookRepair} onChange={(e) => setWebhookRepair(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      {/* 11-12. API 키 (읽기전용) */}
      <Field label="아임웹 API 키" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3 font-mono">
          {maskEnv('IMWEB_API_KEY')} — Vercel 환경변수에서 변경
        </p>
      </Field>

      <Field label="롯데택배 API 키" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3 font-mono">
          {maskEnv('LOTTE_CLIENT_KEY')} — Vercel 환경변수에서 변경
        </p>
      </Field>

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}><Save size={14} />{saving ? '저장 중...' : '저장'}</Button>
      </div>
    </div>
  );
}

function maskEnv(name: string) {
  return `${name}=****` + ' (보안상 마스킹)';
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (<div><label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>{desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}{children}</div>);
}

function Toggle({ checked, onChange }: { checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <button onClick={() => onChange(!checked)}
      className={`relative w-11 h-6 rounded-full transition ${checked ? 'bg-neutral-900' : 'bg-neutral-200'}`}>
      <span className={`absolute top-0.5 left-0.5 w-5 h-5 rounded-full bg-white shadow transition-transform ${checked ? 'translate-x-5' : ''}`} />
    </button>
  );
}
