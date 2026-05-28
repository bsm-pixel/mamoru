'use client';

import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { useContract, useSendContractNotification, useDeleteContract } from '@/hooks/use-contracts';
import { formatKRW, formatDate, formatDateTime } from '@/lib/utils/format';
import { Send, Receipt, Image, ExternalLink, FileText, Trash2 } from 'lucide-react';
import toast from 'react-hot-toast';
import { useState } from 'react';

const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600',
  signed: 'bg-blue-100 text-blue-700',
  sent: 'bg-purple-100 text-purple-700',
  completed: 'bg-green-100 text-green-700',
  cancelled: 'bg-red-100 text-red-700',
};

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중', signed: '서명완료', sent: '발송완료', completed: '완료', cancelled: '취소',
};

const PAYMENT_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합', cms: 'CMS 자동이체',
};

const DELIVERY_LABEL: Record<string, string> = {
  shipping: '본사 발송', pickup: '직접 수령',
};

interface Props {
  contractId: string;
  /** 삭제 후 부모에서 선택 해제할 수 있도록 — 사이드 패널 닫기용 */
  onDeleted?: () => void;
}

export function ContractDetailPanel({ contractId, onDeleted }: Props) {
  const router = useRouter();
  const { data, isLoading } = useContract(contractId);
  const sendNotify = useSendContractNotification();
  const deleteContract = useDeleteContract();
  const [showConvertConfirm, setShowConvertConfirm] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);

  if (isLoading) {
    return <div className="space-y-4"><Skeleton className="h-48" /><Skeleton className="h-32" /></div>;
  }
  if (!data) {
    return (
      <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
        <FileText size={28} className="mb-2 opacity-40" />
        <p className="text-sm">계약서를 찾을 수 없습니다</p>
      </div>
    );
  }

  const { contract, items } = data;

  return (
    <div className="flex-1 overflow-y-auto space-y-4 pr-1">
      {/* 헤더 */}
      <div className="flex items-center justify-between">
        <div>
          <h3 className="text-sm font-bold text-stone-900">{contract.contract_number}</h3>
          <p className="text-xs text-neutral-500 mt-0.5">{contract.customer_name} 님 · {formatDate(contract.created_at)}</p>
        </div>
        <Badge className={STATUS_COLOR[contract.status] || ''}>
          {STATUS_LABEL[contract.status] || contract.status}
        </Badge>
      </div>

      {/* 고객 정보 */}
      <Card>
        <div className="grid grid-cols-2 gap-3 text-sm">
          <div>
            <span className="text-xs text-neutral-500">고객명</span>
            <p className="font-semibold">{contract.customer_name} <span className="text-neutral-400 font-normal">님</span></p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">연락처</span>
            <p>{contract.customer_phone || '-'}</p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">결제방법</span>
            <p>
              {PAYMENT_LABEL[contract.payment_method] || contract.payment_method}
              {contract.installment_months > 0 && ` (${contract.installment_months}개월)`}
            </p>
          </div>
          {contract.customer_email && (
            <div>
              <span className="text-xs text-neutral-500">이메일</span>
              <p className="truncate">{contract.customer_email}</p>
            </div>
          )}
        </div>

        {(contract.customer_title || contract.shop_name) && (
          <div className="mt-3 pt-3 border-t border-neutral-100 grid grid-cols-2 gap-3 text-sm">
            {contract.customer_title && <div><span className="text-xs text-neutral-500">직함</span><p>{contract.customer_title}</p></div>}
            {contract.shop_name && <div><span className="text-xs text-neutral-500">매장명</span><p>{contract.shop_name}</p></div>}
          </div>
        )}

        {contract.delivery_method && (
          <div className="mt-3 pt-3 border-t border-neutral-100 text-sm">
            <span className="text-xs text-neutral-500">수령방법</span>
            <p>{DELIVERY_LABEL[contract.delivery_method] || contract.delivery_method}</p>
          </div>
        )}

        {contract.memo && (
          <p className="mt-3 pt-3 border-t border-neutral-100 text-sm text-neutral-600">{contract.memo}</p>
        )}
      </Card>

      {/* 항목 */}
      <Card>
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">계약 항목</h4>
        <div className="space-y-2">
          {items.map((item) => (
            <div key={item.id} className="flex items-center justify-between py-1.5">
              <div>
                <p className="text-sm font-medium">{item.product_name}</p>
                <p className="text-xs text-neutral-500">
                  {item.sku && !item.sku.startsWith('IW-') && `${item.sku} · `}{formatKRW(item.unit_price)} x {item.quantity}
                </p>
              </div>
              <span className="text-sm font-bold">{formatKRW(item.total_price)}</span>
            </div>
          ))}
        </div>
        <div className="mt-3 pt-3 border-t border-neutral-200 space-y-1">
          {contract.discount_amount > 0 && (
            <div className="flex justify-between text-sm">
              <span className="text-neutral-500">할인</span>
              <span className="text-red-600">-{formatKRW(contract.discount_amount)}</span>
            </div>
          )}
          <div className="flex justify-between text-sm font-bold">
            <span>최종 금액</span>
            <span className="text-stone-900">{formatKRW(contract.final_amount)}</span>
          </div>
          {(contract.deposit_amount > 0 || contract.balance_amount > 0) && (
            <div className="flex justify-between text-xs text-neutral-500 pt-1">
              <span>선납 {formatKRW(contract.deposit_amount)}</span>
              <span>잔금 {formatKRW(contract.balance_amount)}</span>
            </div>
          )}
        </div>
      </Card>

      {/* 서명 */}
      {(contract.signature_data || contract.seller_signature) && (
        <Card>
          <h4 className="text-xs font-semibold text-neutral-500 mb-2">서명</h4>
          <div className="grid grid-cols-2 gap-3">
            {contract.signature_data && (
              <div>
                <p className="text-xs text-neutral-400 mb-1">구매자</p>
                <div className="border border-neutral-200 rounded-lg p-2 bg-white inline-block">
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img src={contract.signature_data} alt="구매자 서명" className="max-w-[160px] h-auto" />
                </div>
              </div>
            )}
            {contract.seller_signature && (
              <div>
                <p className="text-xs text-neutral-400 mb-1">판매자</p>
                <div className="border border-neutral-200 rounded-lg p-2 bg-white inline-block">
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img src={contract.seller_signature} alt="판매자 서명" className="max-w-[160px] h-auto" />
                </div>
              </div>
            )}
          </div>
          {contract.signed_at && (
            <p className="text-xs text-neutral-400 mt-2">서명: {formatDateTime(contract.signed_at)}</p>
          )}
        </Card>
      )}

      {/* 계약서 이미지 */}
      {contract.image_url && (
        <Card>
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <Image size={14} className="text-neutral-500" />
              <span className="text-xs font-semibold text-neutral-500">계약서 이미지</span>
            </div>
            <a href={contract.image_url} target="_blank" rel="noopener noreferrer"
              className="flex items-center gap-1 text-xs text-blue-600 hover:underline">
              열람 <ExternalLink size={12} />
            </a>
          </div>
        </Card>
      )}

      {/* 판매 전환 */}
      {['signed', 'sent', 'draft'].includes(contract.status) && !contract.offline_sale_id && (
        <Card>
          <div className="flex items-center gap-2 mb-2">
            <Receipt size={14} className="text-neutral-600" />
            <span className="text-xs font-semibold text-neutral-600">판매 전환</span>
          </div>
          <div className="flex gap-2">
            <Button
              variant="secondary"
              size="sm"
              className="flex-1"
              onClick={() => setShowConvertConfirm(true)}
            >
              판매 전환
            </Button>
          </div>
        </Card>
      )}

      {/* 판매 전환 확인 */}
      <ConfirmModal
        open={showConvertConfirm}
        onClose={() => setShowConvertConfirm(false)}
        onConfirm={() => {
          router.push(`/sales/new?customer_name=${encodeURIComponent(contract.customer_name)}&customer_phone=${encodeURIComponent(contract.customer_phone || '')}&contract_id=${contract.id}`);
        }}
        title="판매 전환"
        message={`${contract.customer_name}님의 계약을 판매로 전환합니다. 판매 입력 화면으로 이동합니다.`}
        confirmLabel="판매 전환"
      />

      {/* 위험 영역 — 계약서 영구 삭제 (판매 전환된 계약은 disabled) */}
      <Card>
        <div className="flex items-center justify-between gap-2">
          <div className="min-w-0">
            <p className="text-xs font-semibold text-neutral-700">계약서 삭제</p>
            <p className="text-[11px] text-neutral-400 mt-0.5">
              {contract.offline_sale_id
                ? '판매 전환된 계약은 삭제할 수 없습니다'
                : '계약 항목·서명·시리얼 연결이 함께 정리됩니다 (복구 불가)'}
            </p>
          </div>
          <Button
            variant="ghost"
            size="sm"
            className="text-red-500 hover:text-red-600 shrink-0"
            disabled={!!contract.offline_sale_id || deleteContract.isPending}
            onClick={() => setShowDeleteConfirm(true)}
            title={contract.offline_sale_id ? '판매 전환된 계약은 삭제할 수 없습니다' : '계약서 영구 삭제'}
          >
            <Trash2 size={14} />
            {deleteContract.isPending ? '삭제 중...' : '삭제'}
          </Button>
        </div>
      </Card>

      <ConfirmModal
        open={showDeleteConfirm}
        onClose={() => setShowDeleteConfirm(false)}
        onConfirm={async () => {
          await deleteContract.mutateAsync(contract.id);
          onDeleted?.();
        }}
        title="계약서 영구 삭제"
        message={
          <span>
            <strong className="text-neutral-900">{contract.contract_number}</strong> 계약서를 영구 삭제합니다.
            <br />
            계약 항목·서명·시리얼 연결이 모두 정리되며 <strong>복구할 수 없습니다</strong>.
          </span>
        }
        confirmLabel="영구 삭제"
        variant="danger"
      />

      {/* 상세 페이지 링크 */}
      <button
        onClick={() => router.push(`/contracts/${contractId}`)}
        className="w-full text-center text-xs text-neutral-400 hover:text-neutral-600 py-2 transition"
      >
        전체 화면에서 보기 →
      </button>
    </div>
  );
}
