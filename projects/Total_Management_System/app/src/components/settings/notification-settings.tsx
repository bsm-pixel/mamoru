'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save, Bell, Send, Smartphone } from 'lucide-react';
import toast from 'react-hot-toast';
import type { TabProps } from '@/app/(dashboard)/settings/page';
import GoogleCalendarSettings from '@/components/settings/google-calendar-settings';
import BannerSettings from '@/components/settings/banner-settings';
import { requestPushToken } from '@/lib/firebase/client';

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
  const [webhookEvent, setWebhookEvent] = useState('');
  // 🔴 앱 푸시 on/off 토글 제거(2026-08-01) — 고객 행동 푸시는 항상 발송(무조건). 놓치면 안 되므로 게이팅 없음.
  // 067: 후기 요청 자동 발송 정책 토글
  const [reviewAutoRequest, setReviewAutoRequest] = useState(false);
  // 109: 판매 출고 알림톡 (집하 자동감지 시 B2C 고객에게 발송) — 코드엔 있었으나 화면에 토글이 없었음
  const [salesShipped, setSalesShipped] = useState(false);
  // 앱 화면 열려 있을 때 in-app 알림음(notification.wav) — 배송설정 탭에서 여기로 이동(2026-08-01), 기본 ON
  const [soundEnabled, setSoundEnabled] = useState(true);

  useEffect(() => {
    setMasterEnabled(parse(settings['notifications.master_enabled'], true));
    setConsultationReceived(parse(settings['notifications.consultation_received'], true));
    setRepairReceived(parse(settings['notifications.repair_received'], true));
    setRepairCostNotice(parse(settings['notifications.repair_cost_notice'], true));
    setRepairPayment(parse(settings['notifications.repair_payment_confirmed'], true));
    setRepairShipped(parse(settings['notifications.repair_shipped'], true));
    setReviewRequest(parse(settings['notifications.review_request'], true));
    setSalesShipped(parse(settings['notifications.sales_shipped'], false));
    setWebhookConsultation(parse(settings['notifications.webhook_consultation'], ''));
    setWebhookAsReceived(parse(settings['notifications.webhook_as_received'], ''));
    setWebhookRepair(parse(settings['notifications.webhook_repair'], ''));
    setWebhookEvent(parse(settings['notifications.webhook_event'], ''));
    setSoundEnabled(parse(settings['notifications.sound_enabled'], true));
    setReviewAutoRequest(parse(settings['review.auto_request_on_completion'], false));
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
      { key: 'notifications.sales_shipped', value: salesShipped },
      { key: 'notifications.webhook_consultation', value: webhookConsultation },
      { key: 'notifications.webhook_as_received', value: webhookAsReceived },
      { key: 'notifications.webhook_repair', value: webhookRepair },
      { key: 'notifications.webhook_event', value: webhookEvent },
      { key: 'notifications.sound_enabled', value: soundEnabled },
      { key: 'review.auto_request_on_completion', value: reviewAutoRequest },
    ]);
  };

  const NOTIF_ITEMS = [
    { key: 'consultation_received', label: '상담 접수 확인', state: consultationReceived, setter: setConsultationReceived },
    { key: 'repair_received', label: '복원수리 접수 확인', state: repairReceived, setter: setRepairReceived },
    { key: 'repair_cost_notice', label: '비용안내', state: repairCostNotice, setter: setRepairCostNotice },
    { key: 'repair_payment', label: '입금확인', state: repairPayment, setter: setRepairPayment },
    { key: 'repair_shipped', label: '복원수리 출고', state: repairShipped, setter: setRepairShipped },
    { key: 'sales_shipped', label: '판매 출고 안내 (기사님 수거 시 자동)', state: salesShipped, setter: setSalesShipped },
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
      <div className="rounded-lg border border-neutral-200 bg-stone-50 p-4 space-y-3">
        <div>
          <h3 className="text-sm font-bold text-stone-900">📱 내 푸시 알림</h3>
          <p className="text-xs text-neutral-500 mt-0.5">
            고객 접수·행동이 발생하면 사장님 디바이스로 <b>항상</b> 푸시 알림을 보냅니다. (크롬/모바일 앱)
          </p>
          <p className="text-[11px] text-neutral-400 mt-1">
            ※ 놓치면 안 되는 알림이라 on/off 설정 없이 무조건 발송됩니다. 안 오면 아래 <b>테스트</b>로 기기 연결을 확인하세요.
          </p>
        </div>

        {/* 앱 열려 있을 때 알림음 */}
        <div className="flex items-center justify-between py-1 border-t border-neutral-200 pt-3">
          <div>
            <span className="text-sm font-medium">앱 화면 열어둘 때 알림음 🔊</span>
            <p className="text-[11px] text-neutral-400 mt-0.5">
              TMS를 보고 있을 때 접수가 들어오면 &ldquo;띵&rdquo; 소리. (백그라운드/잠금 상태 소리는 휴대폰 알림 소리 설정을 따름)
            </p>
          </div>
          <Toggle checked={soundEnabled} onChange={setSoundEnabled} />
        </div>

        {/* 테스트 발송 패널 */}
        <PushTestPanel />

        {/* 디바이스 정리 — 중복 알림 해결 */}
        <CleanupDevicesPanel />
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

      {/* 067 → 2026-05-25 → 2026-05-26 정책 정정: 약속 ✓ 고객만 자동 발송 (사장님 의도) */}
      <Field
        label="배송완료 시 자동 후기요청"
        desc="복원수리·아임웹 주문·TMS 판매 3채널 공통. OFF → 모든 후기 요청 수동만 (자동 발송 안 됨). ON → 판매 상세에서 '리뷰 약속' 토글 ON 한 건만, ALPS 인수자등록(코드 41/45) 자동 감지 시 알림톡 자동 발송. ※ 약속 OFF 건은 토글 ON 이어도 사장님 수동 발송만. ※ 상담은 정책상 영구 수동만 (2026-04-30)."
      >
        <div className="flex items-center gap-3">
          <Toggle checked={reviewAutoRequest} onChange={setReviewAutoRequest} />
          <span className={`text-sm font-medium ${reviewAutoRequest ? 'text-blue-600' : 'text-neutral-500'}`}>
            {reviewAutoRequest ? 'ON — 약속 ✓ 건 자동 발송' : 'OFF — 모두 수동 발송'}
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

      {/* Make 웹훅 URL — 4개 시나리오 */}
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

      <Field label="Make 웹훅 URL (이벤트)" desc="EVENT 접수확인/입금확인/출고완료 — 비우면 상담 웹훅으로 폴백">
        <input value={webhookEvent} onChange={(e) => setWebhookEvent(e.target.value)}
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

/** 푸시 테스트 패널 — 각 알림 타입별 테스트 발송 */
function PushTestPanel() {
  const [busy, setBusy] = useState<string | null>(null);

  const TESTS: Array<{ type: string; label: string; icon: string }> = [
    { type: 'generic',          label: '기본 테스트 (토글 무관)',      icon: '🔔' },
    { type: 'review',           label: '리뷰 작성',                    icon: '⭐' },
    { type: 'consultation',     label: '상담 접수 (매장방문)',         icon: '📋' },
    { type: 'field_request',    label: '상담 접수 (출장)',             icon: '🚗' },
    { type: 'talk_received',    label: '상담 접수 (톡)',               icon: '💬' },
    { type: 'field_confirmed',  label: '출장 일정 확정',               icon: '✅' },
    { type: 'field_reschedule', label: '출장 일정 재요청',             icon: '🔄' },
    { type: 'repair_received',  label: '복원수리 접수',                icon: '🛠' },
    { type: 'order_received',   label: '아임웹 주문 접수',             icon: '📦' },
  ];

  const fire = async (type: string, label: string) => {
    setBusy(type);
    try {
      const res = await fetch('/api/push/test', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type }),
      });
      const json = await res.json();
      if (!res.ok || !json.ok) {
        toast.error(`테스트 실패: ${json.error || 'unknown'}`);
        return;
      }
      const { sent, failed } = json.data;
      if (sent === 0 && failed === 0) {
        toast(`${label} 테스트 — 등록된 디바이스 없음 (알림 권한 허용 후 새로고침)`, { icon: '⚠️' });
      } else if (sent > 0) {
        toast.success(`${label} 테스트 발송 (${sent}건 성공${failed > 0 ? `, ${failed}건 실패` : ''})`);
      } else {
        toast.error(`${label} 테스트 — ${failed}건 모두 실패 (Vercel 로그 확인)`);
      }
    } catch (err) {
      toast.error(`테스트 실패: ${String(err)}`);
    } finally {
      setBusy(null);
    }
  };

  return (
    <div className="mt-3 pt-3 border-t border-neutral-200 space-y-2">
      <div className="flex items-center gap-1.5">
        <Bell size={13} className="text-neutral-600" />
        <span className="text-xs font-bold text-neutral-700">테스트 발송</span>
        <span className="text-[10px] text-neutral-400 ml-1">※ 실제 등록된 기기로 발송됩니다</span>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5">
        {TESTS.map((t) => (
          <button
            key={t.type}
            onClick={() => fire(t.type, t.label)}
            disabled={busy !== null}
            className="flex items-center gap-1.5 px-2.5 py-1.5 rounded border border-neutral-200 bg-white hover:bg-neutral-50 text-[11px] text-neutral-700 disabled:opacity-50 text-left"
          >
            <span>{t.icon}</span>
            <span className="flex-1 truncate">{t.label}</span>
            {busy === t.type ? (
              <span className="text-[9px] text-neutral-400">...</span>
            ) : (
              <Send size={10} className="text-neutral-400" />
            )}
          </button>
        ))}
      </div>
      <p className="text-[10px] text-neutral-400 leading-relaxed">
        💡 모든 테스트가 실제 등록된 기기로 무조건 발송됩니다 (기기 연결 확인용).
      </p>
    </div>
  );
}

/** 디바이스 정리 패널 — "이 기기만 알림 받기" 버튼
 *  같은 사용자가 여러 토큰 누적되어 푸시 알림이 중복 도착할 때, 한 번 클릭으로
 *  현재 기기 토큰만 남기고 본인의 다른 모든 토큰을 정리.
 */
function CleanupDevicesPanel() {
  const [busy, setBusy] = useState(false);

  const handleCleanup = async () => {
    setBusy(true);
    try {
      const token = await requestPushToken();
      if (!token) {
        toast.error('현재 기기의 토큰을 발급받지 못했습니다. 알림 권한 허용 후 페이지 새로고침해주세요.');
        return;
      }
      const res = await fetch('/api/push/cleanup-others', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ token }),
      });
      const json = await res.json();
      if (!res.ok || !json.ok) {
        toast.error(`정리 실패: ${json.error || 'unknown'}`);
        return;
      }
      const deleted = json.deleted ?? 0;
      if (deleted === 0) {
        toast('정리할 다른 기기 토큰이 없습니다 — 이미 단일 기기 상태입니다', { icon: '✅' });
      } else {
        toast.success(`다른 기기 토큰 ${deleted}건 정리 완료`);
      }
    } catch (err) {
      toast.error(`정리 실패: ${String(err)}`);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="mt-3 pt-3 border-t border-neutral-200 space-y-2">
      <div className="flex items-center gap-1.5">
        <Smartphone size={13} className="text-neutral-600" />
        <span className="text-xs font-bold text-neutral-700">기기 정리</span>
      </div>
      <p className="text-[11px] text-neutral-500 leading-relaxed">
        같은 알림이 여러 번 도착할 때 사용. 현재 이 기기의 토큰만 남기고
        다른 기기·캐시에 등록된 본인의 토큰을 모두 삭제합니다.
      </p>
      <button
        type="button"
        onClick={handleCleanup}
        disabled={busy}
        className="w-full flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg border border-neutral-200 bg-white hover:bg-neutral-50 text-xs font-semibold text-neutral-700 disabled:opacity-50"
      >
        <Smartphone size={13} />
        {busy ? '정리 중...' : '이 기기만 알림 받기'}
      </button>
    </div>
  );
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
