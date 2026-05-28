'use client';

/**
 * Google Calendar 연동 설정 UI
 * 설정 > 알림·연동 탭에 삽입
 *
 * - 연결 버튼 → OAuth 인가 URL로 리다이렉트
 * - 연결 상태 표시 + 이메일 + Workspace/일반 판별
 * - 전체 재동기화 / 연결 해제 버튼
 * - 에러 배너 (refresh_token 만료 등)
 */

import { useEffect, useState } from 'react';
import { Button } from '@/components/ui/button';
import { Calendar, ExternalLink, RefreshCw, Unlink, AlertCircle, CheckCircle2 } from 'lucide-react';
import toast from 'react-hot-toast';

interface StatusData {
  connected: boolean;
  email?: string;
  hd?: string;
  connected_at?: string;
  last_error?: string;
  last_success_at?: string;
}

export default function GoogleCalendarSettings() {
  const [status, setStatus] = useState<StatusData | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);

  const loadStatus = async () => {
    try {
      const res = await fetch('/api/google/calendar/status');
      const json = await res.json();
      if (json.ok) setStatus(json.data as StatusData);
    } catch {
      setStatus(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadStatus();

    // URL 쿼리로 콜백 결과 토스트 표시
    const url = new URL(window.location.href);
    const calParam = url.searchParams.get('google_calendar');
    if (calParam === 'connected') {
      toast.success('Google Calendar 연결 완료');
      url.searchParams.delete('google_calendar');
      window.history.replaceState({}, '', url.toString());
    } else if (calParam === 'error') {
      const msg = url.searchParams.get('msg') || '알 수 없는 오류';
      toast.error(`연결 실패: ${decodeURIComponent(msg)}`);
      url.searchParams.delete('google_calendar');
      url.searchParams.delete('msg');
      window.history.replaceState({}, '', url.toString());
    }
  }, []);

  const handleConnect = () => {
    window.location.href = '/api/google/calendar/auth';
  };

  const handleDisconnect = async () => {
    if (!confirm('Google Calendar 연결을 해제하시겠습니까?\n이미 생성된 이벤트는 그대로 유지됩니다.')) return;
    setBusy(true);
    try {
      const res = await fetch('/api/google/calendar/disconnect', { method: 'POST' });
      const json = await res.json();
      if (json.ok) {
        toast.success('연결 해제됨');
        await loadStatus();
      } else {
        toast.error('해제 실패: ' + (json.error || ''));
      }
    } finally {
      setBusy(false);
    }
  };

  const handleResync = async () => {
    if (!confirm('과거 60일 ~ 미래 180일의 확정/완료 건을 Google Calendar로 일괄 동기화합니다.\n진행하시겠습니까?')) return;
    setBusy(true);
    try {
      const res = await fetch('/api/google/calendar/resync', { method: 'POST' });
      const json = await res.json();
      if (json.ok) {
        const { total, success, failed } = json.data;
        toast.success(`재동기화 완료: ${success}건 성공${failed > 0 ? `, ${failed}건 실패` : ''} (총 ${total}건)`);
        await loadStatus();
      } else {
        toast.error('재동기화 실패: ' + (json.error || ''));
      }
    } finally {
      setBusy(false);
    }
  };

  if (loading) {
    return (
      <div className="rounded-lg border border-neutral-200 bg-stone-50 p-4">
        <div className="text-sm text-neutral-500">Google Calendar 상태 로드 중...</div>
      </div>
    );
  }

  const hasError = !!status?.last_error;
  const isWorkspace = !!status?.hd;

  return (
    <div className="rounded-lg border border-neutral-200 bg-stone-50 p-4 space-y-4">
      {/* 헤더 */}
      <div className="flex items-start gap-3">
        <Calendar size={20} className="text-neutral-700 mt-0.5 shrink-0" />
        <div className="flex-1">
          <h3 className="text-sm font-bold text-stone-900">📅 Google Calendar 연동</h3>
          <p className="text-xs text-neutral-500 mt-0.5">
            상담 확정·변경·취소가 사장님의 Google Calendar에 자동 반영됩니다.
            <br />
            모바일 기본 달력 위젯에서 일정을 바로 확인하실 수 있습니다.
          </p>
        </div>
      </div>

      {/* 에러 배너 */}
      {hasError && (
        <div className="rounded-lg border border-red-200 bg-red-50 p-3 flex items-start gap-2">
          <AlertCircle size={16} className="text-red-600 mt-0.5 shrink-0" />
          <div className="flex-1 text-xs">
            <p className="font-semibold text-red-700">연결 오류</p>
            <p className="text-red-600 mt-1 font-mono break-all">{status?.last_error}</p>
            <p className="text-red-500 mt-1">재연결이 필요할 수 있습니다.</p>
          </div>
        </div>
      )}

      {/* 연결 상태 */}
      {status?.connected ? (
        <div className="space-y-3">
          <div className="rounded-lg bg-white border border-neutral-200 p-3 space-y-2">
            <div className="flex items-center gap-2 text-sm">
              <CheckCircle2 size={16} className="text-green-600" />
              <span className="font-semibold text-green-700">연결됨</span>
            </div>
            <div className="text-xs text-neutral-600 space-y-1 pl-6">
              <div>
                <span className="text-neutral-400">계정:</span>{' '}
                <span className="font-mono">{status.email || '(조회 중)'}</span>
                {isWorkspace && (
                  <span className="ml-2 px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 text-[10px] font-semibold">
                    Workspace
                  </span>
                )}
                {!isWorkspace && status.email && (
                  <span className="ml-2 px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-600 text-[10px] font-semibold">
                    Gmail
                  </span>
                )}
              </div>
              {status.connected_at && (
                <div>
                  <span className="text-neutral-400">연결 일시:</span>{' '}
                  <span>{new Date(status.connected_at).toLocaleString('ko-KR')}</span>
                </div>
              )}
              {status.last_success_at && (
                <div>
                  <span className="text-neutral-400">최근 동기화:</span>{' '}
                  <span>{new Date(status.last_success_at).toLocaleString('ko-KR')}</span>
                </div>
              )}
            </div>
          </div>

          {/* 액션 버튼 */}
          <div className="flex flex-wrap gap-2">
            <Button size="sm" variant="secondary" onClick={handleResync} disabled={busy}>
              <RefreshCw size={13} />
              전체 재동기화
            </Button>
            <Button size="sm" variant="ghost" onClick={handleDisconnect} disabled={busy}>
              <Unlink size={13} />
              연결 해제
            </Button>
          </div>

          <p className="text-[11px] text-neutral-400 leading-relaxed">
            💡 매장방문·출장 확정 시 자동으로 캘린더에 등록됩니다.
            취소 시 이벤트 삭제, 완료 시 ✅ 표시 유지, 재요청 시 ⏳ 표시로 변경됩니다.
            <br />
            ⚠️ 캘린더에서 직접 일정을 수정하지 마세요 — TMS와 불일치가 발생할 수 있습니다.
          </p>
        </div>
      ) : (
        <div className="space-y-3">
          <div className="rounded-lg bg-white border border-neutral-200 p-4 text-center">
            <p className="text-sm text-neutral-600 mb-3">
              Google Calendar에 연결하면
              <br />
              확정된 상담이 자동으로 등록됩니다
            </p>
            <Button size="sm" onClick={handleConnect} disabled={busy}>
              <ExternalLink size={13} />
              Google 계정 연결
            </Button>
          </div>
          <p className="text-[11px] text-neutral-400 leading-relaxed">
            💡 권장 계정: <span className="font-semibold">bsm@mamoru.kr</span>
            <br />
            연결 시 Google 로그인 화면으로 이동합니다. 권한 승인 후 자동으로 돌아옵니다.
          </p>
        </div>
      )}
    </div>
  );
}
