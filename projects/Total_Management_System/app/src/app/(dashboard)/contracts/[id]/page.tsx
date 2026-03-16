'use client';

import { use } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useContract, useSendContractNotification } from '@/hooks/use-contracts';
import { formatKRW, formatDate, formatDateTime } from '@/lib/utils/format';
import { ArrowLeft, Send, Receipt } from 'lucide-react';

const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600',
  signed: 'bg-blue-100 text-blue-700',
  sent: 'bg-purple-100 text-purple-700',
  completed: 'bg-green-100 text-green-700',
  cancelled: 'bg-red-100 text-red-700',
};

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중',
  signed: '서명완료',
  sent: '발송완료',
  completed: '완료',
  cancelled: '취소',
};

const PAYMENT_LABEL: Record<string, string> = {
  card: '카드',
  cash: '현금',
  transfer: '계좌이체',
  mixed: '복합',
  cms: 'CMS 자동이체',
};

const DELIVERY_LABEL: Record<string, string> = {
  shipping: '본사 발송',
  pickup: '직접 수령',
};

export default function ContractDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();
  const { data, isLoading } = useContract(id);
  const sendNotify = useSendContractNotification();

  if (isLoading) {
    return (
      <>
        <Topbar title="계약서 상세" />
        <div className="px-4 md:px-6 py-4 space-y-4">
          <Skeleton className="h-48" />
          <Skeleton className="h-32" />
        </div>
      </>
    );
  }

  if (!data) {
    return (
      <>
        <Topbar title="계약서 상세" />
        <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
          계약서를 찾을 수 없습니다
        </div>
      </>
    );
  }

  const { contract, items } = data;

  return (
    <>
      <Topbar title="계약서 상세" />

      <div className="px-4 md:px-6 py-4 space-y-4 max-w-3xl">
        <Button variant="ghost" size="sm" onClick={() => router.push('/contracts')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>

        {/* 계약 정보 */}
        <Card>
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-bold text-indigo-black">{contract.contract_number}</h3>
            <Badge className={STATUS_COLOR[contract.status] || ''}>
              {STATUS_LABEL[contract.status] || contract.status}
            </Badge>
          </div>

          <div className="grid grid-cols-2 gap-3 text-sm">
            <div>
              <span className="text-xs text-neutral-500">고객명</span>
              <p className="font-semibold">{contract.customer_name}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">연락처</span>
              <p>{contract.customer_phone || '-'}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">이메일</span>
              <p>{contract.customer_email || '-'}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">주소</span>
              <p className="truncate">{contract.customer_address || '-'}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">작성일</span>
              <p>{formatDateTime(contract.created_at)}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">결제방법</span>
              <p>
                {PAYMENT_LABEL[contract.payment_method] || contract.payment_method}
                {contract.installment_months > 0 && ` (${contract.installment_months}개월)`}
              </p>
            </div>
          </div>

          {/* 확장 필드: 매장 정보 */}
          {(contract.customer_title || contract.shop_name || contract.shop_address) && (
            <div className="mt-3 pt-3 border-t border-neutral-100 grid grid-cols-2 gap-3 text-sm">
              {contract.customer_title && (
                <div>
                  <span className="text-xs text-neutral-500">직함</span>
                  <p>{contract.customer_title}</p>
                </div>
              )}
              {contract.shop_name && (
                <div>
                  <span className="text-xs text-neutral-500">매장명</span>
                  <p>{contract.shop_name}</p>
                </div>
              )}
              {contract.shop_address && (
                <div className="col-span-2">
                  <span className="text-xs text-neutral-500">매장주소</span>
                  <p>{contract.shop_address}</p>
                </div>
              )}
            </div>
          )}

          {/* 수령방법 */}
          {contract.delivery_method && (
            <div className="mt-3 pt-3 border-t border-neutral-100 text-sm">
              <span className="text-xs text-neutral-500">수령방법</span>
              <p>
                {DELIVERY_LABEL[contract.delivery_method] || contract.delivery_method}
                {contract.unavailable_days && ` (불가: ${contract.unavailable_days})`}
              </p>
            </div>
          )}

          {/* 선납/잔금 */}
          {(contract.deposit_amount > 0 || contract.balance_amount > 0) && (
            <div className="mt-3 pt-3 border-t border-neutral-100 grid grid-cols-2 gap-3 text-sm">
              <div>
                <span className="text-xs text-neutral-500">선납금</span>
                <p className="font-semibold">{formatKRW(contract.deposit_amount)}</p>
              </div>
              <div>
                <span className="text-xs text-neutral-500">잔금</span>
                <p className="font-semibold">{formatKRW(contract.balance_amount)}</p>
              </div>
            </div>
          )}

          {contract.memo && (
            <p className="mt-3 pt-3 border-t border-neutral-100 text-sm text-neutral-600">
              {contract.memo}
            </p>
          )}
        </Card>

        {/* 항목 */}
        <Card>
          <h3 className="text-sm font-bold text-indigo-black mb-3">계약 항목</h3>
          <div className="space-y-2">
            {items.map((item) => (
              <div key={item.id} className="flex items-center justify-between py-2 border-b border-neutral-50 last:border-0">
                <div>
                  <p className="text-sm font-medium">{item.product_name}</p>
                  <p className="text-xs text-neutral-500">
                    {item.sku && `${item.sku} · `}{formatKRW(item.unit_price)} x {item.quantity}
                    {item.option_text && ` · ${item.option_text}`}
                  </p>
                </div>
                <span className="text-sm font-bold">{formatKRW(item.total_price)}</span>
              </div>
            ))}
          </div>

          <div className="mt-3 pt-3 border-t border-neutral-200 space-y-1">
            <div className="flex justify-between text-sm">
              <span className="text-neutral-500">소계</span>
              <span>{formatKRW(contract.total_amount)}</span>
            </div>
            {contract.discount_amount > 0 && (
              <div className="flex justify-between text-sm">
                <span className="text-neutral-500">할인</span>
                <span className="text-red-600">-{formatKRW(contract.discount_amount)}</span>
              </div>
            )}
            <div className="flex justify-between text-sm font-bold">
              <span>최종 금액</span>
              <span className="text-terracotta">{formatKRW(contract.final_amount)}</span>
            </div>
          </div>
        </Card>

        {/* 서명 */}
        {(contract.signature_data || contract.seller_signature) && (
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3">서명</h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {contract.signature_data && (
                <div>
                  <p className="text-xs text-neutral-500 mb-1">구매자</p>
                  <div className="border border-neutral-200 rounded-lg p-2 bg-white inline-block">
                    {/* eslint-disable-next-line @next/next/no-img-element */}
                    <img src={contract.signature_data} alt="구매자 서명" className="max-w-[200px] h-auto" />
                  </div>
                </div>
              )}
              {contract.seller_signature && (
                <div>
                  <p className="text-xs text-neutral-500 mb-1">판매자</p>
                  <div className="border border-neutral-200 rounded-lg p-2 bg-white inline-block">
                    {/* eslint-disable-next-line @next/next/no-img-element */}
                    <img src={contract.seller_signature} alt="판매자 서명" className="max-w-[200px] h-auto" />
                  </div>
                </div>
              )}
            </div>
            {contract.signed_at && (
              <p className="text-xs text-neutral-500 mt-2">
                서명 시각: {formatDateTime(contract.signed_at)}
              </p>
            )}
          </Card>
        )}

        {/* 알림톡 발송 */}
        {contract.status === 'signed' && !contract.notification_sent_at && (
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-2">알림톡 발송</h3>
            <p className="text-xs text-neutral-500 mb-3">
              고객에게 계약서 사본을 알림톡으로 발송합니다.
            </p>
            <Button
              variant="secondary"
              size="sm"
              onClick={() => sendNotify.mutate(contract.id)}
              disabled={sendNotify.isPending || !contract.customer_phone}
            >
              <Send size={14} />
              {sendNotify.isPending ? '발송 중...' : '알림톡 발송'}
            </Button>
          </Card>
        )}

        {/* 다음 단계: 판매 등록 CTA */}
        {(contract.status === 'signed' || contract.status === 'sent') && !contract.offline_sale_id && (
          <Card className="border-green-200 bg-green-50/50">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Receipt size={18} className="text-green-600" />
                <div>
                  <p className="text-sm font-semibold text-green-700">판매 등록</p>
                  <p className="text-xs text-neutral-500">계약 완료 → 판매를 등록합니다</p>
                </div>
              </div>
              <Button
                size="sm"
                onClick={() => router.push(`/sales/new?customer_name=${encodeURIComponent(contract.customer_name)}&customer_phone=${encodeURIComponent(contract.customer_phone || '')}&contract_id=${contract.id}`)}
              >
                등록하기
              </Button>
            </div>
          </Card>
        )}

        {contract.notification_sent_at && (
          <p className="text-xs text-neutral-500">
            알림톡 발송 완료: {formatDateTime(contract.notification_sent_at)}
          </p>
        )}
      </div>
    </>
  );
}
