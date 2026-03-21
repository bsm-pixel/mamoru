'use client';

import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { useOrderSync, useProductSync } from '@/hooks/use-orders';
import { createClient } from '@/lib/supabase/client';
import { formatDateTime } from '@/lib/utils/format';
import { RefreshCw, Database, CheckCircle2, AlertCircle, Package, Upload } from 'lucide-react';
import type { SyncLog } from '@/lib/supabase/types';

export default function SettingsPage() {
  const supabase = createClient();
  const sync = useOrderSync();
  const productSync = useProductSync();

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

        {/* 아임웹 상품 동기화 */}
        <Card>
          <CardHeader>
            <CardTitle>아임웹 상품 동기화</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-500 mb-4">
            아임웹에 등록된 상품을 TMS에 동기화합니다.
            매입가·거래처는 TMS에서 별도 관리됩니다.
          </p>
          <Button
            onClick={() => productSync.mutate()}
            disabled={productSync.isPending}
          >
            <Package size={16} className={productSync.isPending ? 'animate-spin' : ''} />
            {productSync.isPending ? '동기화 중...' : '상품 동기화'}
          </Button>
          {productSync.data && (
            <p className="text-xs text-neutral-500 mt-2">
              {productSync.data.created}개 생성, {productSync.data.updated}개 업데이트
              {productSync.data.errors?.length > 0 && ` (오류 ${productSync.data.errors.length}건)`}
            </p>
          )}
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
                        {log.sync_type === 'imweb_products' ? '상품' : '주문'} {log.records_synced}건
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

        {/* 이카운트 데이터 이관 */}
        <Card>
          <CardHeader>
            <CardTitle>이카운트 데이터 이관</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-500 mb-4">
            이카운트에서 추출한 엑셀(CSV)을 업로드하여 고객·판매내역을 TMS로 이관합니다.
          </p>
          <div className="space-y-3">
            <CsvUploadButton
              label="고객 CSV 업로드"
              endpoint="/api/import/customers"
              resultLabel={(d) => `${d.created}명 등록, ${d.skipped}명 스킵`}
            />
            <CsvUploadButton
              label="판매내역 CSV 업로드"
              endpoint="/api/import/sales"
              resultLabel={(d) => `${d.sales_created}건 판매, ${d.serials_created}개 시리얼 등록`}
            />
          </div>
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

/** CSV 파일 업로드 + API 호출 버튼 */
function CsvUploadButton({ label, endpoint, resultLabel }: {
  label: string;
  endpoint: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  resultLabel: (data: any) => string;
}) {
  const [uploading, setUploading] = useState(false);
  const [result, setResult] = useState<string | null>(null);

  async function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploading(true);
    setResult(null);

    try {
      const text = await file.text();
      const lines = text.split('\n').filter((l) => l.trim());
      if (lines.length < 2) throw new Error('데이터가 없습니다');

      // CSV 파싱 (첫 줄 = 헤더)
      const headers = lines[0].split(',').map((h) => h.trim().replace(/"/g, ''));
      const rows = lines.slice(1).map((line) => {
        const values = line.split(',').map((v) => v.trim().replace(/"/g, ''));
        const row: Record<string, string> = {};
        headers.forEach((h, i) => { row[h] = values[i] || ''; });
        return row;
      });

      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows }),
      });

      if (!res.ok) throw new Error((await res.json()).error);
      const data = await res.json();
      setResult(resultLabel(data));
    } catch (err) {
      setResult(`오류: ${String(err)}`);
    } finally {
      setUploading(false);
      e.target.value = '';
    }
  }

  return (
    <div>
      <label className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-neutral-100 text-sm font-medium text-neutral-700 hover:bg-neutral-200 transition cursor-pointer">
        <Upload size={14} />
        {uploading ? '업로드 중...' : label}
        <input type="file" accept=".csv,.txt" onChange={handleFile} disabled={uploading} className="hidden" />
      </label>
      {result && (
        <p className={`text-xs mt-1.5 ${result.startsWith('오류') ? 'text-red-500' : 'text-green-600'}`}>
          {result}
        </p>
      )}
    </div>
  );
}
