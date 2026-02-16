'use client';

import { useQuery } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { useOrderSync } from '@/hooks/use-orders';
import { createClient } from '@/lib/supabase/client';
import { formatDateTime } from '@/lib/utils/format';
import { RefreshCw, Database, CheckCircle2, AlertCircle } from 'lucide-react';
import type { SyncLog } from '@/lib/supabase/types';

export default function SettingsPage() {
  const supabase = createClient();
  const sync = useOrderSync();

  const { data: syncLogs } = useQuery({
    queryKey: ['sync-logs'],
    queryFn: async () => {
      const { data } = await supabase
        .from('sync_log')
        .select('*')
        .order('started_at', { ascending: false })
        .limit(10);
      return (data || []) as SyncLog[];
    },
  });

  return (
    <>
      <Topbar title="설정" />

      <div className="px-4 md:px-6 py-4 space-y-6 max-w-2xl">
        {/* 아임웹 동기화 */}
        <Card>
          <CardHeader>
            <CardTitle>아임웹 주문 동기화</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-500 mb-4">
            아임웹에서 주문 데이터를 가져와 TMS에 동기화합니다.
            자동 동기화는 Vercel Cron으로 5분마다 실행됩니다.
          </p>
          <Button
            onClick={() => sync.mutate()}
            disabled={sync.isPending}
          >
            <RefreshCw size={16} className={sync.isPending ? 'animate-spin' : ''} />
            {sync.isPending ? '동기화 중...' : '지금 동기화'}
          </Button>
        </Card>

        {/* 동기화 이력 */}
        <Card>
          <CardHeader>
            <CardTitle>동기화 이력</CardTitle>
          </CardHeader>
          {!syncLogs || syncLogs.length === 0 ? (
            <div className="flex items-center justify-center h-20 text-sm text-neutral-400">
              <Database size={16} className="mr-2" />
              동기화 이력이 없습니다
            </div>
          ) : (
            <div className="divide-y divide-neutral-100 -mx-5">
              {syncLogs.map((log) => (
                <div key={log.id} className="flex items-center justify-between px-5 py-3">
                  <div>
                    <div className="flex items-center gap-2">
                      {log.status === 'completed' ? (
                        <CheckCircle2 size={14} className="text-success" />
                      ) : log.status === 'failed' ? (
                        <AlertCircle size={14} className="text-error" />
                      ) : (
                        <RefreshCw size={14} className="text-info animate-spin" />
                      )}
                      <span className="text-sm font-medium">
                        {log.records_synced}건 동기화
                      </span>
                      <Badge variant={log.status === 'completed' ? 'success' : log.status === 'failed' ? 'error' : 'info'}>
                        {log.status === 'completed' ? '완료' : log.status === 'failed' ? '실패' : '진행중'}
                      </Badge>
                    </div>
                    {log.error_message && (
                      <p className="text-xs text-error mt-1 truncate max-w-xs">
                        {log.error_message}
                      </p>
                    )}
                  </div>
                  <span className="text-xs text-neutral-400">
                    {formatDateTime(log.started_at)}
                  </span>
                </div>
              ))}
            </div>
          )}
        </Card>

        {/* 환경변수 안내 */}
        <Card>
          <CardHeader>
            <CardTitle>환경변수 설정</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-500 mb-3">
            아래 환경변수가 <code className="text-xs bg-warm-ivory px-1 py-0.5 rounded">.env.local</code>에 설정되어 있어야 합니다.
          </p>
          <div className="bg-indigo-black rounded-lg p-4 text-xs font-mono text-cream/80 space-y-1 overflow-x-auto">
            <p>NEXT_PUBLIC_SUPABASE_URL=</p>
            <p>NEXT_PUBLIC_SUPABASE_ANON_KEY=</p>
            <p>SUPABASE_SERVICE_ROLE_KEY=</p>
            <p>IMWEB_API_KEY=</p>
            <p>IMWEB_API_SECRET=</p>
            <p>LOTTE_API_URL=</p>
            <p>LOTTE_CANCEL_API_URL=</p>
            <p>LOTTE_CLIENT_KEY=</p>
            <p>LOTTE_JOBCUSTCD=</p>
            <p>LOTTE_SENDER_NAME=</p>
            <p>LOTTE_SENDER_TEL=</p>
            <p>LOTTE_SENDER_ZIP=</p>
            <p>LOTTE_SENDER_ADDR=</p>
            <p>CRON_SECRET=</p>
          </div>
        </Card>
      </div>
    </>
  );
}
