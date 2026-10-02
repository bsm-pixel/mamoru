/**
 * 리뷰 이벤트 당첨자 배송 — 공용 로직 (2026-10-02, 마이그 156)
 *
 * 당첨 마킹 SSOT = reviews.event_month / event_rank (룰렛·직접지정 화면이 그대로 씀).
 * review_event_shipments 는 그 당첨자들의 "배송"만 담당 — 목록을 열 때 당첨자 기준으로 자동 동기화한다.
 *
 * 흐름도: projects/Total_Management_System/docs/TMS_FLOW_REVIEW_EVENT.md 「당첨자 배송」
 */

import crypto from 'crypto';
import { sendNotification } from '@/lib/notification/make-webhook';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export const ADDRESS_PAGE = 'https://page.mamoru.kr/projects/reviews/page_event_address.html';

export interface ShipmentRow {
  id: string;
  review_id: string;
  event_month: string;
  rank: number | null;
  rank_label: string | null;
  prize: string | null;
  name: string;
  phone: string;
  token: string;
  postcode: string | null;
  address_road: string | null;
  address_detail: string | null;
  delivery_message: string | null;
  address_prefill_source: string | null;
  address_submitted_at: string | null;
  customer_id: string | null;
  won_notified_at: string | null;
  invoice_number: string | null;
  courier_name: string | null;
  invoice_created_at: string | null;
  shipped_at: string | null;
  shipped_notified_at: string | null;
  delivered_at: string | null;
  cancelled_at: string | null;
  memo: string | null;
  created_at: string;
}

interface Prize { rank: number; label?: string; name?: string; count?: number }

export const digits = (v: string | null | undefined) => String(v || '').replace(/[^0-9]/g, '');

export function addressLink(token: string): string {
  return `${ADDRESS_PAGE}?t=${token}`;
}

function prizeInfo(prizes: Prize[] | null | undefined, rank: number | null): { label: string; name: string } {
  const p = (prizes || []).find((x) => Number(x.rank) === Number(rank));
  const label = (p?.label || '').trim() || (rank ? `${rank}등` : '당첨');
  const name = (p?.name || '').trim() || label;
  return { label, name };
}

/**
 * 해당 월 당첨자 기준으로 배송 행을 맞춘다.
 *  - 새 당첨자 → 행 생성 (토큰 발급)
 *  - 송장 전인 행 → 등수·상품명을 최신 설정으로 갱신 (상품명 나중에 바꿔도 반영)
 *  - 당첨에서 빠진 사람 → 송장 전이면 cancelled_at (soft). 송장 이후면 그대로 둔다(실물 이동 중)
 */
export async function syncShipments(db: Db, month: string): Promise<ShipmentRow[]> {
  const [{ data: winners }, { data: cfg }, { data: existing }] = await Promise.all([
    db.from('reviews').select('id, name, phone, event_rank').eq('event_month', month).not('event_rank', 'is', null),
    db.from('review_event_config').select('prizes').eq('month', month).maybeSingle(),
    db.from('review_event_shipments').select('*').eq('event_month', month),
  ]);
  const prizes: Prize[] = cfg?.prizes || [];
  const byReview = new Map<string, ShipmentRow>((existing || []).map((r: ShipmentRow) => [r.review_id, r]));
  const winnerIds = new Set<string>();

  for (const w of winners || []) {
    winnerIds.add(w.id);
    const { label, name } = prizeInfo(prizes, w.event_rank);
    const row = byReview.get(w.id);
    if (!row) {
      await db.from('review_event_shipments').insert({
        review_id: w.id,
        event_month: month,
        rank: w.event_rank,
        rank_label: label,
        prize: name,
        name: (w.name || '').trim() || '고객',
        phone: digits(w.phone),
        token: crypto.randomBytes(16).toString('hex'),
      });
    } else if (!row.invoice_number && (row.rank !== w.event_rank || row.prize !== name || row.rank_label !== label || row.cancelled_at)) {
      await db.from('review_event_shipments')
        .update({ rank: w.event_rank, rank_label: label, prize: name, cancelled_at: null, updated_at: new Date().toISOString() })
        .eq('id', row.id);
    }
  }
  // 당첨에서 빠진 사람 — 송장 전이면 취소 표시
  for (const r of existing || []) {
    if (!winnerIds.has(r.review_id) && !r.invoice_number && !r.cancelled_at) {
      await db.from('review_event_shipments').update({ cancelled_at: new Date().toISOString() }).eq('id', r.id);
    }
  }

  const { data: rows } = await db.from('review_event_shipments')
    .select('*').eq('event_month', month).is('cancelled_at', null)
    .order('rank', { ascending: true }).order('created_at', { ascending: true });
  return rows || [];
}

/**
 * 이 전화번호로 예전에 남긴 주소 — 고객정보 → 복원수리 접수 → 아임웹 주문 순 (가장 최근·정확한 것부터)
 */
export async function findPrefillAddress(db: Db, phone: string): Promise<{
  postcode: string; address_road: string; address_detail: string; source: 'customer' | 'repair' | 'order'; customer_id: string | null;
} | null> {
  const p = digits(phone);
  if (!p) return null;

  const { data: cust } = await db.from('customers')
    .select('id, postcode, address_road, address_detail')
    .eq('phone_normalized', p).is('merged_into_id', null)
    .order('updated_at', { ascending: false }).limit(1);
  const c = cust?.[0];
  if (c?.address_road) {
    return { postcode: c.postcode || '', address_road: c.address_road, address_detail: c.address_detail || '', source: 'customer', customer_id: c.id };
  }

  const { data: rep } = await db.from('repairs')
    .select('postcode, address, address_detail')
    .eq('phone_normalized', p).not('address', 'is', null)
    .order('created_at', { ascending: false }).limit(1);
  if (rep?.[0]?.address) {
    return { postcode: rep[0].postcode || '', address_road: rep[0].address, address_detail: rep[0].address_detail || '', source: 'repair', customer_id: c?.id ?? null };
  }

  const { data: ord } = await db.from('orders')
    .select('recipient_postcode, recipient_address, recipient_address_detail')
    .or(`recipient_phone.eq.${p},orderer_phone.eq.${p}`).not('recipient_address', 'is', null)
    .order('ordered_at', { ascending: false }).limit(1);
  if (ord?.[0]?.recipient_address) {
    return { postcode: ord[0].recipient_postcode || '', address_road: ord[0].recipient_address, address_detail: ord[0].recipient_address_detail || '', source: 'order', customer_id: c?.id ?? null };
  }
  return null;
}

/**
 * 고객이 저장한 주소를 TMS 고객정보에도 반영 (사장님 결정 2026-10-02: 저장된 주소 = 고객의 신규 주소)
 *  - 같은 전화의 고객이 있으면 주소 갱신, 없으면 고객 신규 생성 (전화 = 고객 매칭 키)
 */
export async function upsertCustomerAddress(db: Db, s: { name: string; phone: string; postcode: string; address_road: string; address_detail: string }): Promise<string | null> {
  const p = digits(s.phone);
  const now = new Date().toISOString();
  const { data: cust } = await db.from('customers')
    .select('id').eq('phone_normalized', p).is('merged_into_id', null)
    .order('updated_at', { ascending: false }).limit(1);
  if (cust?.[0]) {
    await db.from('customers')
      .update({ postcode: s.postcode, address_road: s.address_road, address_detail: s.address_detail || null, updated_at: now })
      .eq('id', cust[0].id);
    return cust[0].id;
  }
  const { data: created, error } = await db.from('customers')
    // source 는 기존 4종(imweb/consultation/as/manual) 안에서 — 출처는 memo 로 남긴다
    .insert({ name: s.name, phone: p, phone_normalized: p, postcode: s.postcode, address_road: s.address_road, address_detail: s.address_detail || null, source: 'manual', memo: '리뷰 이벤트 당첨 배송지 입력으로 등록' })
    .select('id').single();
  if (error) {
    console.warn('[event-shipments] 고객 생성 실패(주소는 배송 행에만 저장):', error.message);
    return null;
  }
  return created?.id ?? null;
}

/** 당첨 안내 알림톡 (배송지 입력 링크 포함) */
export async function sendWonNotice(db: Db, row: ShipmentRow): Promise<{ ok: boolean; error?: string }> {
  const r = await sendNotification({
    template: 'review_event_won',
    phone: row.phone,
    name: row.name,
    data: {
      id: row.id,
      rank: row.rank_label || (row.rank ? `${row.rank}등` : ''),
      prize: row.prize || '',
      token: row.token,
      address_link: addressLink(row.token),
    },
  });
  if (r.success) {
    await db.from('review_event_shipments').update({ won_notified_at: new Date().toISOString() }).eq('id', row.id);
  }
  return { ok: r.success, error: r.error };
}

/** 당첨 상품 출고 알림톡 (집하 감지 시) */
export async function sendShippedNotice(db: Db, row: Pick<ShipmentRow, 'id' | 'name' | 'phone' | 'prize' | 'invoice_number' | 'courier_name'>): Promise<{ ok: boolean; error?: string }> {
  const r = await sendNotification({
    template: 'review_event_shipped',
    phone: row.phone,
    name: row.name,
    data: {
      id: row.id,
      prize: row.prize || '',
      tracking: row.invoice_number || '',
      courier: row.courier_name || '롯데택배',
    },
  });
  if (r.success) {
    await db.from('review_event_shipments').update({ shipped_notified_at: new Date().toISOString() }).eq('id', row.id);
  }
  return { ok: r.success, error: r.error };
}
