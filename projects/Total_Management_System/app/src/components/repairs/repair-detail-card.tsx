'use client';

import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { formatPhone, formatKRW, formatDate } from '@/lib/utils/format';
import type { Repair } from '@/lib/supabase/types';
import { User, MapPin, Scissors } from 'lucide-react';

interface RepairDetailCardProps {
  repair: Repair;
}

export function RepairDetailCard({ repair: r }: RepairDetailCardProps) {
  return (
    <div className="space-y-4">
      {/* 고객 정보 */}
      <Card>
        <CardHeader>
          <CardTitle>
            <User size={16} className="inline mr-1.5" />
            고객 정보
          </CardTitle>
        </CardHeader>
        <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
          <dt className="text-neutral-500">이름</dt>
          <dd className="font-medium">{r.name}</dd>
          <dt className="text-neutral-500">전화</dt>
          <dd>{formatPhone(r.phone)}</dd>
          <dt className="text-neutral-500">진행방식</dt>
          <dd>
            {r.proceed_type ? (
              <span className="inline-block px-2 py-0.5 rounded-full text-xs font-medium bg-info-soft text-info">
                {r.proceed_type}
              </span>
            ) : '-'}
          </dd>
          <dt className="text-neutral-500">전달방법</dt>
          <dd>{r.delivery_method || '-'}</dd>
        </dl>
      </Card>

      {/* 주소 */}
      {(r.address || r.postcode) && (
        <Card>
          <CardHeader>
            <CardTitle>
              <MapPin size={16} className="inline mr-1.5" />
              주소
            </CardTitle>
          </CardHeader>
          <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
            <dt className="text-neutral-500">우편번호</dt>
            <dd>{r.postcode || '-'}</dd>
            <dt className="text-neutral-500">주소</dt>
            <dd className="col-span-2">{r.address || '-'} {r.address_detail || ''}</dd>
          </dl>
        </Card>
      )}

      {/* 수량 + 날짜 */}
      <Card>
        <CardHeader>
          <CardTitle>
            <Scissors size={16} className="inline mr-1.5" />
            접수 정보
          </CardTitle>
        </CardHeader>
        <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
          <dt className="text-neutral-500">마모루</dt>
          <dd className="font-medium">{r.qty_mamoru}자루</dd>
          <dt className="text-neutral-500">타사</dt>
          <dd className="font-medium">{r.qty_other}자루</dd>
          <dt className="text-neutral-500">수거요청일</dt>
          <dd>{r.pickup_date ? formatDate(r.pickup_date, 'yyyy.MM.dd') : '-'}</dd>
          <dt className="text-neutral-500">접수일시</dt>
          <dd>{r.received_at ? formatDate(r.received_at, 'yyyy.MM.dd HH:mm') : '-'}</dd>
        </dl>
      </Card>

      {/* 비용 */}
      <Card>
        <CardHeader>
          <CardTitle>비용 정보</CardTitle>
        </CardHeader>
        <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
          <dt className="text-neutral-500">수리비</dt>
          <dd className="font-medium">{formatKRW(r.service_cost)}</dd>
          <dt className="text-neutral-500">수거비</dt>
          <dd>{formatKRW(r.shipping_fee)}</dd>
          <dt className="text-neutral-500 font-semibold">합계</dt>
          <dd className="font-bold text-terracotta-deep">{formatKRW(r.total_amount)}</dd>
        </dl>
      </Card>

      {/* 메모 */}
      {r.memo && (
        <Card>
          <CardHeader>
            <CardTitle>고객 메모</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-700 whitespace-pre-wrap">{r.memo}</p>
        </Card>
      )}

      {/* 관리자 메모 */}
      {r.admin_note && (
        <Card>
          <CardHeader>
            <CardTitle>관리자 메모 (추가전달사항)</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-700 whitespace-pre-wrap">{r.admin_note}</p>
        </Card>
      )}
    </div>
  );
}
