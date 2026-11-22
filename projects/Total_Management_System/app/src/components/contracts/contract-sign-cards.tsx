'use client';

/**
 * 계약서 상세의 「고객 서명 링크」·「이니셜 각인」 칸 (2026-11-22, 마이그 172)
 *  사이드 패널(contract-detail-panel)과 전체 화면(contracts/[id]) 이 함께 쓴다.
 */

import { useState } from 'react';
import toast from 'react-hot-toast';
import { Link2, Copy, PenLine } from 'lucide-react';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useCreateSignLink, useUpdateEngraving } from '@/hooks/use-contracts';
import { signMessage, signUrl, ENGRAVING_MAX } from '@/lib/contracts/sign-link';
import { ENGRAVING_NOTICE } from '@/lib/contracts/notice';
import { formatDateTime } from '@/lib/utils/format';
import { TONE } from '@/lib/ui/status-tone';
import type { Contract } from '@/lib/supabase/types';

async function copy(text: string, done: string) {
  try { await navigator.clipboard.writeText(text); toast.success(done); }
  catch { toast.error('복사하지 못했어요 — 길게 눌러 직접 복사해 주세요'); }
}

export function ContractSignCards({ contract }: { contract: Contract }) {
  const makeLink = useCreateSignLink();
  const saveEng = useUpdateEngraving();
  const [engEdit, setEngEdit] = useState<string | null>(null);

  const signed = !!contract.signature_data;
  const cancelled = contract.status === 'cancelled';
  const token = contract.sign_token;
  const engPending = contract.engraving && contract.engraving_status === 'pending';

  return (
    <>
      {/* 고객 서명 링크 — 서명 전이거나, 서명 뒤에도 각인 문구를 링크로 받을 때 */}
      {!cancelled && (!signed || (engPending && token)) && (
        <Card data-sign-link-card>
          <div className="flex items-center gap-2 mb-2">
            <Link2 size={14} className="text-neutral-600" />
            <span className="text-xs font-semibold text-neutral-600">{signed ? '각인 문구 받는 링크' : '고객 서명 링크'}</span>
            {token && contract.sign_link_at && <span className="text-[11px] text-neutral-400 ml-auto">만든 때 {formatDateTime(contract.sign_link_at)}</span>}
          </div>
          {!token ? (
            <>
              <p className="text-xs text-neutral-500 mb-2">링크를 톡으로 보내면 고객이 폰에서 계약 내용을 읽고 서명합니다. 서명되면 휴대폰으로 알려 드려요.</p>
              <Button size="sm" className="w-full" disabled={makeLink.isPending} onClick={() => makeLink.mutate(contract.id)} data-make-sign-link>
                {makeLink.isPending ? '만드는 중…' : '서명 링크 만들기'}
              </Button>
            </>
          ) : (
            <>
              <p className="text-[11px] text-neutral-500 break-all mb-2 font-mono">{signUrl(token)}</p>
              <div className="grid grid-cols-2 gap-2">
                <Button size="sm" onClick={() => copy(signMessage(contract.customer_name, token, { engravingOnly: signed }), '톡에 붙여 넣을 문구를 복사했어요')} data-copy-sign-message>
                  <Copy size={13} /> 보낼 문구 복사
                </Button>
                <Button size="sm" variant="secondary" onClick={() => copy(signUrl(token), '주소를 복사했어요')}>주소만 복사</Button>
              </div>
              {!signed && <p className="text-[11px] text-neutral-400 mt-2">고객이 서명하면 이 칸이 「서명」으로 바뀝니다.</p>}
            </>
          )}
        </Card>
      )}

      {/* 이니셜 각인 */}
      {contract.engraving && (
        <Card data-engraving-card>
          <div className="flex items-center gap-2 mb-2">
            <PenLine size={14} className="text-neutral-600" />
            <span className="text-xs font-semibold text-neutral-600">이니셜 각인</span>
            <span className={`ml-auto text-[11px] font-semibold px-2 py-0.5 rounded-full ${engPending ? TONE.wait : TONE.done}`}>
              {engPending ? '문구 기다리는 중' : '문구 받음'}
            </span>
          </div>
          {engEdit === null ? (
            <div className="flex items-center gap-2">
              <p className={`flex-1 min-w-0 text-base font-bold tracking-wide break-all ${engPending ? 'text-neutral-400 font-normal text-sm' : 'text-neutral-900'}`}>
                {contract.engraving_text || '아직 받지 않았어요'}
              </p>
              <Button size="sm" variant="secondary" onClick={() => setEngEdit(contract.engraving_text || '')}>{engPending ? '받은 문구 적기' : '고치기'}</Button>
            </div>
          ) : (
            <div className="space-y-2">
              <input
                autoFocus value={engEdit} maxLength={ENGRAVING_MAX} onChange={(e) => setEngEdit(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') saveEng.mutate({ id: contract.id, engraving_text: engEdit }, { onSuccess: () => setEngEdit(null) }); if (e.key === 'Escape') setEngEdit(null); }}
                placeholder="예: S.M.B" className="w-full border border-neutral-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:border-neutral-800"
                data-engraving-input
              />
              <div className="flex gap-2">
                <Button size="sm" className="flex-1" disabled={saveEng.isPending} onClick={() => saveEng.mutate({ id: contract.id, engraving_text: engEdit }, { onSuccess: () => setEngEdit(null) })}>저장</Button>
                <Button size="sm" variant="ghost" onClick={() => setEngEdit(null)}>닫기</Button>
              </div>
              <p className="text-[11px] text-neutral-400">비우고 저장하면 다시 「문구 기다리는 중」이 됩니다.</p>
            </div>
          )}
          {contract.engraving_at && !engPending && <p className="text-[11px] text-neutral-400 mt-2">받은 때 {formatDateTime(contract.engraving_at)}</p>}
          <p className="text-[11px] text-neutral-500 leading-relaxed mt-3 pt-2 border-t border-neutral-100">{ENGRAVING_NOTICE}</p>
        </Card>
      )}
    </>
  );
}
