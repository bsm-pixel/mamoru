/**
 * 판매 출고 알림톡(sales_shipped) 발송 — 수동/자동 공용 (109, 2026-07-12)
 *
 * 진입점 2개가 같은 내용을 보내야 한다:
 *   1) 수동  — api/sales/[id] PATCH action='mark_shipped' (사장님 버튼 + 알림톡 체크)
 *   2) 자동  — api/cron/track-delivery 집하 감지 블록 [3-A] (롯데 기사 수거 스캔)
 * 이 함수로 묶지 않으면 두 경로의 알림톡 내용이 조용히 갈라진다.
 *
 * ⚠️ B2B(딜러·아카데미) 는 발송 대상이 아니다 — 호출부에서 isB2BCustomerType() 으로 걸러야 한다.
 *    (여기서도 2차 가드를 둔다: customerType 을 넘기면 B2B 는 skip)
 */
import { sendNotification } from '@/lib/notification/make-webhook';
import { isB2BCustomerType } from '@/lib/sales/customer-type';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface SalesShippedTarget {
  id: string;                       // offline_sales.id (품명 조회용)
  saleNumber: string | null;
  invoiceNumber: string | null;
  customerName: string;
  customerPhone: string | null;
  customerType?: string | null;     // B2B 2차 가드
  courierName?: string | null;
}

/** 품명 조립 — "블런트 5.5인치 ×2, 틴닝 28발" 형태 */
async function buildGoodsName(db: Db, saleId: string): Promise<string> {
  const { data: items } = await db
    .from('offline_sale_items')
    .select('product_name, quantity')
    .eq('sale_id', saleId);

  if (!items || items.length === 0) return '마모루 제품';
  return items
    .map((i: { product_name: string; quantity: number }) =>
      i.quantity > 1 ? `${i.product_name} ×${i.quantity}` : i.product_name,
    )
    .join(', ');
}

/**
 * EVENT 접수페이지 → 판매전환 건 판별.
 * convertEventToSale 이 memo 를 'EVENT 전환 (…)' 로만 남긴다(LS='재고판매 전환', 수동=사용자입력).
 * ∴ 품목 category='EVENT' 가 아니라 memo 접두어로 봐야 EVENT 카탈로그 제품의 '수동 판매' 오발송을 막는다.
 */
async function isEventOriginSale(db: Db, saleId: string): Promise<boolean> {
  const { data } = await db.from('offline_sales').select('memo').eq('id', saleId).single();
  return typeof data?.memo === 'string' && data.memo.startsWith('EVENT 전환');
}

/**
 * 출고 알림톡 발송.
 * 발송 조건(전화번호·B2B)을 스스로 검사하므로 호출부는 결과만 보면 된다.
 * 반환: { sent: true } 면 실제로 발송 성공 → 호출부가 shipped_notified_at 기록.
 */
export async function sendSalesShippedNotification(
  db: Db,
  sale: SalesShippedTarget,
): Promise<{ sent: boolean; reason?: string; error?: string }> {
  if (!sale.customerPhone) return { sent: false, reason: 'no_phone' };
  if (isB2BCustomerType(sale.customerType)) return { sent: false, reason: 'b2b' };  // 2차 가드

  const goodsName = await buildGoodsName(db, sale.id);
  // EVENT 접수페이지 유입 건은 전용 출고완료(event_shipped)로 발송 → EVENT 시나리오로 분리
  const isEvent = await isEventOriginSale(db, sale.id);
  const template = isEvent ? 'event_shipped' : 'sales_shipped';

  // 토글(notifications.sales_shipped / event_shipped) 체크는 sendNotification 내부에서 수행 → 여기서 중복 확인 불필요
  const result = await sendNotification({
    template,
    phone: sale.customerPhone,
    name: sale.customerName,
    data: {
      id: sale.saleNumber || sale.id,
      tracking: sale.invoiceNumber || '',
      courier: sale.courierName || '롯데택배',
      goods_name: goodsName,
    },
  });

  // 🔴 토글 OFF 로 건너뛴 건 '발송'이 아니다 — shipped_notified_at 을 찍으면 '발송됨'으로 잘못 뜬다
  if (result.skipped) return { sent: false, reason: 'toggle_off' };
  if (!result.success) return { sent: false, reason: 'send_failed', error: result.error };
  return { sent: true };
}
