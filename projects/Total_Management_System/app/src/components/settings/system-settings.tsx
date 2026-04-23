'use client';

import { useState, useEffect } from 'react';
import { useQuery } from '@tanstack/react-query';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Save, RefreshCw, CheckCircle2, AlertCircle, Database, Upload } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';
import { useOrderSync, useProductSync } from '@/hooks/use-orders';
import { formatDateTime } from '@/lib/utils/format';
import type { TabProps } from '@/app/(dashboard)/settings/page';
import type { SyncLog } from '@/lib/supabase/types';
import { NAV_GROUPS } from '@/lib/utils/constants';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

export default function SystemSettings({ settings, onSave, saving }: TabProps) {
  const supabase = createClient();
  const sync = useOrderSync();
  const productSync = useProductSync();
  const [pageSize, setPageSize] = useState(20);
  const [startPage, setStartPage] = useState('/dashboard');
  const [sidebarConfig, setSidebarConfig] = useState<{ order: string[]; hidden: string[] }>({ order: [], hidden: [] });

  const { data: syncLogs } = useQuery({
    queryKey: ['sync-logs'],
    queryFn: async () => {
      const { data } = await supabase.from('sync_log').select('*').order('started_at', { ascending: false }).limit(10);
      return (data || []) as SyncLog[];
    },
  });

  useEffect(() => {
    setPageSize(parse(settings['system.table_page_size'], 20));
    setStartPage(parse(settings['system.start_page'], '/dashboard'));
    setSidebarConfig(parse(settings['system.sidebar_config'], { order: [], hidden: [] }));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'system.table_page_size', value: pageSize },
      { key: 'system.start_page', value: startPage },
      { key: 'system.sidebar_config', value: sidebarConfig },
    ]);
  };

  // 사이드바 전체 메뉴 항목 추출
  const allMenuItems = NAV_GROUPS.flatMap((g) => g.items.map((item) => ({
    href: item.href, label: item.label, group: g.group,
  })));

  const toggleHidden = (href: string) => {
    const next = sidebarConfig.hidden.includes(href)
      ? sidebarConfig.hidden.filter((h) => h !== href)
      : [...sidebarConfig.hidden, href];
    setSidebarConfig({ ...sidebarConfig, hidden: next });
  };

  // 시작 페이지 옵션
  const START_PAGES = [
    { value: '/dashboard', label: '대시보드' },
    { value: '/sales', label: '판매 조회' },
    { value: '/orders/dashboard', label: '주문관리' },
    { value: '/repairs', label: '복원수리' },
    { value: '/consultations', label: '상담관리' },
  ];

  // CSV 업로드 (기존 이카운트 이관은 삭제됨)
  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">시스템 설정</h2>

      {/* 1. 사업장 정보 — 회계/판매와 공유 안내 */}
      <Field label="사업장 정보" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3">
          사업자 정보는 <span className="font-semibold">회계 설정</span> 또는 <span className="font-semibold">판매 설정</span>에서 입력하면 전체 시스템에 반영됩니다.
        </p>
      </Field>

      {/* 5. 테이블 행 수 */}
      <Field label="테이블 기본 행 수" desc="모든 목록 페이지의 기본 표시 건수.">
        <select value={pageSize} onChange={(e) => setPageSize(Number(e.target.value))}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          {[10, 20, 25, 50, 100].map((n) => <option key={n} value={n}>{n}건</option>)}
        </select>
      </Field>

      {/* 6. 기본 시작 페이지 */}
      <Field label="기본 시작 페이지" desc="로그인 후 처음 보이는 페이지.">
        <select value={startPage} onChange={(e) => setStartPage(e.target.value)}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          {START_PAGES.map((p) => <option key={p.value} value={p.value}>{p.label}</option>)}
        </select>
      </Field>

      {/* 7. 시스템 버전 */}
      <Field label="시스템 버전" desc="">
        <p className="text-sm bg-neutral-50 rounded-lg px-3 py-2 font-mono">
          TMS v2.5.0 — 설정 탭 리뉴얼
        </p>
      </Field>

      {/* 14. 사이드바 커스텀 */}
      <Field label="사이드바 메뉴 커스텀" desc="메뉴를 숨기거나 표시할 수 있습니다. 숨겨진 메뉴는 URL 직접 접근은 가능합니다.">
        <div className="space-y-1">
          {allMenuItems.map((item) => (
            <label key={item.href} className="flex items-center gap-2 py-1">
              <input type="checkbox"
                checked={!sidebarConfig.hidden.includes(item.href)}
                onChange={() => toggleHidden(item.href)}
                className="rounded" />
              <span className="text-sm">{item.label}</span>
              {item.group && <span className="text-xs text-neutral-400">({item.group})</span>}
            </label>
          ))}
        </div>
      </Field>

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}><Save size={14} />{saving ? '저장 중...' : '저장'}</Button>
      </div>

      {/* ── 기존 기능: 동기화 ── */}
      <div className="pt-6 border-t border-neutral-200">
        <h3 className="text-base font-bold mb-4">아임웹 동기화</h3>

        <div className="flex gap-3 mb-4">
          <Button onClick={() => sync.mutate()} disabled={sync.isPending} variant="secondary">
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            {sync.isPending ? '동기화 중...' : '주문 동기화'}
          </Button>
          <Button onClick={() => productSync.mutate()} disabled={productSync.isPending} variant="secondary">
            <RefreshCw size={14} className={productSync.isPending ? 'animate-spin' : ''} />
            {productSync.isPending ? '동기화 중...' : '상품 동기화'}
          </Button>
        </div>

        {productSync.data && (
          <div className="mb-3 rounded-lg border border-neutral-200 bg-neutral-50 p-3 text-xs space-y-1.5">
            <div className="flex items-center gap-2">
              <span className="font-semibold text-neutral-700">아임웹 API:</span>
              <span className="text-neutral-600">{productSync.data.total_fetched ?? productSync.data.synced}건 조회</span>
              <span className="text-neutral-400">→</span>
              <span className="font-semibold text-green-700">{productSync.data.synced}건 반영</span>
              {productSync.data.errors?.length > 0 && (
                <>
                  <span className="text-neutral-400">·</span>
                  <span className="font-semibold text-red-600">{productSync.data.errors.length}건 실패</span>
                </>
              )}
            </div>
            <div className="text-neutral-500">
              신규 생성 <b className="text-neutral-700">{productSync.data.created}</b> ·
              기존 업데이트 <b className="text-neutral-700">{productSync.data.updated}</b>
              {typeof productSync.data.linked === 'number' && productSync.data.linked > 0 && (
                <> · 수동등록 연결 <b className="text-neutral-700">{productSync.data.linked}</b></>
              )}
            </div>
            {productSync.data.errors?.length > 0 && (
              <details className="pt-1">
                <summary className="cursor-pointer text-red-600 font-semibold hover:underline">
                  실패 내역 {productSync.data.errors.length}건 보기
                </summary>
                <div className="mt-2 bg-white border border-red-200 rounded p-2 max-h-48 overflow-y-auto font-mono text-[10px] text-neutral-700 space-y-1">
                  {productSync.data.errors.slice(0, 50).map((err: string, i: number) => (
                    <div key={i} className="break-all">{err}</div>
                  ))}
                  {productSync.data.errors.length > 50 && (
                    <div className="text-neutral-400">... {productSync.data.errors.length - 50}건 더 (Vercel 로그에서 전체 확인)</div>
                  )}
                </div>
              </details>
            )}
          </div>
        )}

        {/* 동기화 이력 */}
        <h4 className="text-sm font-semibold mb-2">동기화 이력</h4>
        {!syncLogs || syncLogs.length === 0 ? (
          <div className="flex items-center justify-center h-16 text-sm text-neutral-400">
            <Database size={16} className="mr-2" /> 이력 없음
          </div>
        ) : (
          <div className="divide-y divide-neutral-100 rounded-lg border border-neutral-100">
            {syncLogs.map((log) => (
              <div key={log.id} className="flex items-center justify-between px-3 py-2">
                <div className="flex items-center gap-2">
                  {log.status === 'completed' ? <CheckCircle2 size={14} className="text-success" />
                    : log.status === 'failed' ? <AlertCircle size={14} className="text-error" />
                    : <RefreshCw size={14} className="text-info animate-spin" />}
                  <span className="text-xs font-medium">
                    {log.sync_type === 'imweb_products' ? '상품' : '주문'} {log.records_synced}건
                  </span>
                  <Badge variant={log.status === 'completed' ? 'success' : log.status === 'failed' ? 'error' : 'info'}>
                    {log.status === 'completed' ? '완료' : log.status === 'failed' ? '실패' : '진행중'}
                  </Badge>
                </div>
                <span className="text-xs text-neutral-400">{formatDateTime(log.started_at)}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* 환경변수 */}
      <div className="pt-4 border-t border-neutral-200">
        <h3 className="text-base font-bold mb-2">환경변수</h3>
        <div className="bg-indigo-black rounded-lg p-4 text-xs font-mono text-cream/80 space-y-0.5 overflow-x-auto">
          {['NEXT_PUBLIC_SUPABASE_URL', 'NEXT_PUBLIC_SUPABASE_ANON_KEY', 'SUPABASE_SERVICE_ROLE_KEY',
            'IMWEB_API_KEY', 'IMWEB_API_SECRET', 'LOTTE_API_URL', 'LOTTE_CLIENT_KEY',
            'LOTTE_SENDER_NAME', 'LOTTE_SENDER_TEL', 'LOTTE_SENDER_ZIP', 'LOTTE_SENDER_ADDR',
            'MAKE_WEBHOOK_URL', 'MAKE_REPAIR_WEBHOOK_URL', 'CRON_SECRET',
          ].map((env) => <p key={env}>{env}=****</p>)}
        </div>
      </div>
    </div>
  );
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (<div><label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>{desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}{children}</div>);
}
