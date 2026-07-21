import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';
import { sendSalesShippedNotification } from '@/lib/notification/sales-shipped';
import { sendReviewRequestNotification } from '@/lib/notification/review-request';
import { getServerSetting } from '@/hooks/use-settings';
import { recalcOutstanding } from '@/lib/outstanding';

/** PATCH /api/sales/[id] — 취소 / 결제상태 변경 / 메모 수정 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { action } = body as { action: string };

    // 현재 판매 건 조회
    const { data: sale, error: saleErr } = await db
      .from('offline_sales')
      .select('*')
      .eq('id', id)
      .single();
    if (saleErr || !sale) {
      return NextResponse.json({ error: '판매 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // --- A) 판매 취소 ---
    if (action === 'cancel') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '이미 취소된 판매입니다' }, { status: 400 });
      }

      const reason = body.reason || '';

      // 1. 시리얼 복원 — previous_zone으로 원래 위치 복원
      const { data: serials } = await db
        .from('product_serials')
        .select('id, product_id, previous_zone')
        .eq('offline_sale_id', id);

      if (serials && serials.length > 0) {
        for (const serial of serials) {
          await db
            .from('product_serials')
            .update({
              status: 'in_stock',
              warehouse_zone: serial.previous_zone || 'ready', // 원래 zone 복원 (NULL이면 ready fallback)
              previous_zone: null,
              sold_via: null,
              offline_sale_id: null,
              sold_at: null,
              sold_to_name: null,
              sold_to_phone: null,
            })
            .eq('id', serial.id);
        }
      }

      // 2. 재고 복원 + 3. 아임웹 재고 동기화
      const { data: items } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (items && items.length > 0) {
        // 상품별 수량 집계 후 병렬 처리
        const productQtyMap: Record<string, number> = {};
        for (const item of items) {
          if (item.product_id && item.quantity) {
            productQtyMap[item.product_id] = (productQtyMap[item.product_id] || 0) + item.quantity;
          }
        }

        await Promise.all(Object.entries(productQtyMap).map(async ([productId, qty]) => {
          const { data: prod } = await db
            .from('products')
            .select('stock_quantity, raw_stock, imweb_product_no')
            .eq('id', productId)
            .single();
          if (!prod) return;

          // 판매 시 raw_stock 는 '시리얼 없는 수량'만큼만 차감됐다 → 취소 시 딱 그만큼만 복원.
          // ⚠️ 시리얼 수는 위 복원 루프가 offline_sale_id 를 null 로 지우기 '전'에 가져온 serials 에서 센다.
          //    (지운 뒤 offline_sale_id 로 세면 항상 0 → 시리얼 판매인데 raw 과다복원 → 무결성 -1 버그. 2026-07-21 fix)
          const serialQty = (serials || []).filter((s: { product_id: string }) => s.product_id === productId).length;
          const rawRestore = Math.max(0, qty - serialQty);

          const newStock = (prod.stock_quantity || 0) + qty;
          const updateData: Record<string, unknown> = { stock_quantity: newStock };
          // 시리얼 없이 팔린 수량(rawRestore)만큼만 raw_stock 복원 (시리얼분은 위에서 in_stock 로 되살림)
          if (rawRestore > 0) {
            updateData.raw_stock = (prod.raw_stock || 0) + rawRestore;
          }
          await db.from('products').update(updateData).eq('id', productId);

          if (prod.imweb_product_no) {
            try {
              await updateImwebStock(Number(prod.imweb_product_no), +qty); // 취소 → 재고 복구(+)
            } catch (e) {
              console.error('[imweb] 취소 재고 동기화 실패:', prod.imweb_product_no, e);
            }
          }
        }));
      }

      // 4. 미수금: 아래 취소 반영(cancelled_at) 후 recalcOutstanding 로 일괄 재계산

      // 5. 판매 상태 업데이트
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({
          cancelled_at: new Date().toISOString(),
          cancelled_reason: reason,
          cancelled_by: user.id,
        })
        .eq('id', id);

      if (updateErr) throw updateErr;

      // 077: 회계 자동 연동 — 입금된 금액이 있으면 환불(expense) 기록
      if (sale.paid_amount > 0) {
        try {
          await db.from('cash_transactions').insert({
            transaction_date: new Date().toISOString().slice(0, 10),
            type: 'expense',
            category: '매출취소환불',
            amount: sale.paid_amount,
            memo: `${sale.sale_number} 취소: ${reason || '사유 미기재'}`,
            source_type: 'offline_sale',
            source_id: id,
            created_by: user.id,
          });
        } catch (e) {
          console.error('[sales cancel] cashflow refund insert 실패:', e);
        }
      }

      await recalcOutstanding(db, sale.customer_id);
      return NextResponse.json({ success: true, action: 'cancelled' });
    }

    // --- A-2) 반품 처리 ---
    if (action === 'return') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 반품 처리할 수 없습니다' }, { status: 400 });
      }
      if (sale.returned_at) {
        return NextResponse.json({ error: '이미 반품 처리된 판매입니다' }, { status: 400 });
      }

      const reason = body.reason || '';

      // 1. 시리얼 복원 — previous_zone으로 원래 위치 복원
      const { data: serials } = await db
        .from('product_serials')
        .select('id, product_id, previous_zone')
        .eq('offline_sale_id', id);

      if (serials && serials.length > 0) {
        for (const serial of serials) {
          await db
            .from('product_serials')
            .update({
              status: 'in_stock',
              warehouse_zone: serial.previous_zone || 'ready',
              previous_zone: null,
              sold_via: null,
              offline_sale_id: null,
              sold_at: null,
              sold_to_name: null,
              sold_to_phone: null,
            })
            .eq('id', serial.id);
        }
      }

      // 2. 재고 복원 + 3. 아임웹 재고 동기화
      const { data: items } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (items && items.length > 0) {
        const productQtyMap: Record<string, number> = {};
        for (const item of items) {
          if (item.product_id && item.quantity) {
            productQtyMap[item.product_id] = (productQtyMap[item.product_id] || 0) + item.quantity;
          }
        }

        await Promise.all(Object.entries(productQtyMap).map(async ([productId, qty]) => {
          const { data: prod } = await db
            .from('products')
            .select('stock_quantity, raw_stock, imweb_product_no')
            .eq('id', productId)
            .single();
          if (!prod) return;

          // (취소와 동일 fix) 시리얼 판매분은 raw_stock 을 안 건드렸으므로, 시리얼 없는 수량만큼만 복원.
          // serials 는 offline_sale_id 를 null 로 지우기 전 값 → 여기서 상품별 시리얼 수를 센다.
          const serialQty = (serials || []).filter((s: { product_id: string }) => s.product_id === productId).length;
          const rawRestore = Math.max(0, qty - serialQty);

          const newStock = (prod.stock_quantity || 0) + qty;
          const updateData: Record<string, unknown> = { stock_quantity: newStock };
          if (rawRestore > 0) {
            updateData.raw_stock = (prod.raw_stock || 0) + rawRestore;
          }
          await db.from('products').update(updateData).eq('id', productId);

          if (prod.imweb_product_no) {
            try {
              await updateImwebStock(Number(prod.imweb_product_no), +qty); // 반품 → 재고 복구(+)
            } catch (e) {
              console.error('[imweb] 반품 재고 동기화 실패:', prod.imweb_product_no, e);
            }
          }
        }));
      }

      // 4. 미수금: 아래 반품 반영(returned_at) 후 recalcOutstanding 로 일괄 재계산

      // 5. 반품 상태 기록
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({
          returned_at: new Date().toISOString(),
          return_reason: reason,
        })
        .eq('id', id);

      if (updateErr) throw updateErr;

      // 077: 회계 자동 연동 — 입금된 금액이 있으면 환불(expense) 기록
      if (sale.paid_amount > 0) {
        try {
          await db.from('cash_transactions').insert({
            transaction_date: new Date().toISOString().slice(0, 10),
            type: 'expense',
            category: '매출반품환불',
            amount: sale.paid_amount,
            memo: `${sale.sale_number} 반품: ${reason || '사유 미기재'}`,
            source_type: 'offline_sale',
            source_id: id,
            created_by: user.id,
          });
        } catch (e) {
          console.error('[sales return] cashflow refund insert 실패:', e);
        }
      }

      await recalcOutstanding(db, sale.customer_id);
      return NextResponse.json({ success: true, action: 'returned' });
    }

    // --- B) 결제상태 변경 ---
    if (action === 'update_payment') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 변경할 수 없습니다' }, { status: 400 });
      }

      const { payment_status, paid_amount } = body as {
        payment_status: string;
        paid_amount?: number;
      };

      if (!payment_status) {
        return NextResponse.json({ error: 'payment_status 필수' }, { status: 400 });
      }

      // 미수금: 아래 결제상태 반영 후 recalcOutstanding 로 일괄 재계산 (±diff 누적 X)
      const updateData: Record<string, unknown> = { payment_status };
      if (paid_amount !== undefined) updateData.paid_amount = paid_amount;
      // 완납 처리 시 paid_amount 미지정이면 총액으로 채워 데이터 정합 유지
      else if (payment_status === 'paid') updateData.paid_amount = sale.total_amount;

      const { error: updateErr } = await db
        .from('offline_sales')
        .update(updateData)
        .eq('id', id);

      if (updateErr) throw updateErr;
      await recalcOutstanding(db, sale.customer_id);
      return NextResponse.json({ success: true, action: 'payment_updated' });
    }

    // --- C) 메모 수정 ---
    if (action === 'update_memo') {
      const { memo } = body as { memo: string };
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({ memo: memo || null })
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'memo_updated' });
    }

    // --- C-2) 포장완료(준비완료) 토글 — 2026-07-18
    //  '내가 물리적으로 포장까지 끝냈다'는 표시. 송장 유무와 무관(포장은 송장 전에도 함).
    //  복원수리(repairs.packed_at)와 동일한 개념 · 동일 컬럼명으로 통일. 알림톡/외부연동 없음(내부 표시 전용).
    if (action === 'mark_packed' || action === 'unmark_packed') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매입니다' }, { status: 400 });
      }
      if (sale.shipped_at) {
        return NextResponse.json({ error: '이미 출고된 건은 변경할 수 없습니다' }, { status: 400 });
      }
      const packedAt = action === 'mark_packed' ? new Date().toISOString() : null;
      const { error: packErr } = await db
        .from('offline_sales')
        .update({ packed_at: packedAt })
        .eq('id', id);

      if (packErr) throw packErr;
      return NextResponse.json({ success: true, action: action, packed_at: packedAt });
    }

    // --- D) 출고완료 처리 ---
    if (action === 'mark_shipped') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매입니다' }, { status: 400 });
      }
      if (!sale.invoice_number) {
        return NextResponse.json({ error: '송장이 없습니다. 먼저 송장을 생성해주세요.' }, { status: 400 });
      }
      if (sale.shipped_at) {
        return NextResponse.json({ error: '이미 출고완료 처리되었습니다' }, { status: 400 });
      }

      const { send_notification } = body as { send_notification?: boolean };

      // 109: 조건부 CAS — 크론(집하 자동감지)이 먼저 처리했으면 여기서 덮어쓰지 않는다
      const { data: claimed, error: updateErr } = await db
        .from('offline_sales')
        .update({ shipped_at: new Date().toISOString(), shipped_source: 'manual' })
        .eq('id', id)
        .is('shipped_at', null)
        .select('id');

      if (updateErr) throw updateErr;
      if (!claimed || claimed.length === 0) {
        return NextResponse.json({ error: '이미 출고완료 처리되었습니다 (기사님 수거 자동 감지)' }, { status: 400 });
      }

      // 선택적 알림톡 발송 (백그라운드) — 109: 자동 경로와 같은 공유 함수 사용
      if (send_notification) {
        after(async () => {
          try {
            const sent = await sendSalesShippedNotification(db, {
              id,
              saleNumber: sale.sale_number,
              invoiceNumber: sale.invoice_number,
              customerName: sale.customer_name,
              customerPhone: sale.customer_phone,
              customerType: sale.customer_type,
              courierName: sale.courier_name,
            });
            if (sent.sent) {
              await db.from('offline_sales')
                .update({ shipped_notified_at: new Date().toISOString() })
                .eq('id', id);
              console.log('[sales mark_shipped notify] 출고 알림톡 발송 성공');
            } else {
              console.log(`[sales mark_shipped notify] 미발송 — ${sent.reason} ${sent.error || ''}`);
            }
          } catch (e) {
            console.error('[sales mark_shipped notify] 예외:', e);
          }
        });
      }

      return NextResponse.json({ success: true, action: 'marked_shipped' });
    }

    // --- D-1) 출고 알림톡 수동 재발송 (2026-07-15) ---
    //   이미 출고됐는데 알림톡이 안 나간 건(토글 OFF 시절 자동 출고 등) 사장님이 직접 보낼 때.
    if (action === 'resend_ship_notify') {
      if (sale.cancelled_at) return NextResponse.json({ error: '취소된 판매입니다' }, { status: 400 });
      if (!sale.shipped_at) return NextResponse.json({ error: '아직 출고되지 않았습니다' }, { status: 400 });
      if (sale.shipped_notified_at) return NextResponse.json({ error: '이미 출고 알림톡이 발송되었습니다' }, { status: 400 });

      const sent = await sendSalesShippedNotification(db, {
        id,
        saleNumber: sale.sale_number,
        invoiceNumber: sale.invoice_number,
        customerName: sale.customer_name,
        customerPhone: sale.customer_phone,
        customerType: sale.customer_type,
        courierName: sale.courier_name,
      });
      if (!sent.sent) {
        // 실패 사유를 그대로 전달 (toggle_off → 설정 안내 / no_phone / b2b / send_failed)
        const msg = sent.reason === 'toggle_off'
          ? '알림톡 설정(판매 출고 안내)이 꺼져 있습니다. 설정에서 켠 뒤 다시 시도하세요.'
          : sent.reason === 'no_phone' ? '고객 전화번호가 없습니다.'
          : sent.reason === 'b2b' ? '거래처(B2B)는 출고 알림톡 대상이 아닙니다.'
          : (sent.error || '발송에 실패했습니다.');
        return NextResponse.json({ error: msg }, { status: 400 });
      }
      await db.from('offline_sales').update({ shipped_notified_at: new Date().toISOString() }).eq('id', id);
      return NextResponse.json({ success: true, action: 'ship_notify_sent' });
    }

    // --- D-2) 배송완료/고객수령 처리 (2026-05-25 Phase 2) ---
    //   mode='delivery': 송장 있는 택배 발송 — ALPS 추적 실패 fallback (수동 배송완료)
    //   mode='pickup':   송장 없는 매장 직접 수령 — "고객 수령 완료"
    if (action === 'mark_delivered') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매입니다' }, { status: 400 });
      }
      if (sale.delivered_at) {
        return NextResponse.json({ error: '이미 배송완료 처리되었습니다' }, { status: 400 });
      }

      const { mode } = body as { mode?: 'delivery' | 'pickup' };
      if (mode !== 'delivery' && mode !== 'pickup') {
        return NextResponse.json({ error: 'invalid_mode' }, { status: 400 });
      }

      // delivery 모드: 송장 있어야 함 (ALPS 추적 fallback 케이스)
      if (mode === 'delivery' && !sale.invoice_number) {
        return NextResponse.json({ error: '송장이 없습니다. 매장 수령 모드로 처리하세요.' }, { status: 400 });
      }
      // pickup 모드: 송장 없어야 함 (매장 직접 수령 — 송장 있으면 delivery 모드로)
      if (mode === 'pickup' && sale.invoice_number) {
        return NextResponse.json({ error: '송장이 있는 건은 배송완료(delivery) 모드로 처리하세요.' }, { status: 400 });
      }

      const now = new Date().toISOString();
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({ delivered_at: now })
        .eq('id', id);

      if (updateErr) throw updateErr;

      // 자동 후기요청 — 택배(ALPS cron)와 동일 기준을 픽업/수동배송완료에도 적용 (2026-06-12)
      //   약속✓ + 토글ON + 미발송 + 연락처 있음 → 즉시 발송 (버튼 클릭=수령완료라 cron 불필요)
      after(async () => {
        try {
          if (sale.review_requested_at || !sale.review_promised_at || !sale.customer_phone) return;
          const autoEnabled = await getServerSetting<boolean>(db, 'review.auto_request_on_completion', false);
          if (!autoEnabled) return;
          const reviewType = (sale.review_promised_type as 'purchase' | 'repair' | 'consult' | null) || 'purchase';
          const subtype = reviewType === 'purchase' ? undefined : (sale.review_promised_subtype as string | null) || undefined;
          const r = await sendReviewRequestNotification({
            source: 'sale',
            sourceId: sale.sale_number,
            customerName: sale.customer_name || '고객',
            customerPhone: sale.customer_phone,
            reviewType,
            subtype,
          });
          if (r.success && !r.skipped) {
            await db.from('offline_sales').update({ review_requested_at: new Date().toISOString() }).eq('id', id);
            console.log(`[mark_delivered/${mode} auto-review] ${sale.sale_number} 발송 성공`);
          }
        } catch (e) {
          console.error('[mark_delivered auto-review] 예외:', e);
        }
      });

      return NextResponse.json({
        success: true,
        action: 'marked_delivered',
        mode,
        delivered_at: now,
      });
    }

    // --- E) 판매 정보 수정 (금액/할인/결제방법/날짜) ---
    if (action === 'edit_sale') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 수정할 수 없습니다' }, { status: 400 });
      }

      const { total_amount, discount_amount, payment_method, sale_date, payment_detail } = body as {
        total_amount?: number;
        discount_amount?: number;
        payment_method?: string;
        sale_date?: string;
        payment_detail?: Record<string, number>;
      };

      const updateData: Record<string, unknown> = {};
      if (total_amount !== undefined) updateData.total_amount = total_amount;
      if (discount_amount !== undefined) updateData.discount_amount = discount_amount;
      if (payment_method !== undefined) updateData.payment_method = payment_method;
      if (sale_date !== undefined) updateData.sale_date = sale_date;
      if (payment_detail !== undefined) updateData.payment_detail = payment_detail;

      // VAT 재계산 (카드 금액 변경 시)
      if (total_amount !== undefined || payment_method !== undefined) {
        const newTotal = total_amount ?? sale.total_amount;
        const newMethod = payment_method ?? sale.payment_method;
        const cardAmount = newMethod === 'card' ? newTotal
          : newMethod === 'mixed' && payment_detail?.card ? payment_detail.card
          : 0;
        if (cardAmount > 0) {
          updateData.supply_amount = Math.round(cardAmount / 1.1);
          updateData.vat_amount = cardAmount - (updateData.supply_amount as number);
          updateData.is_vat_included = true;
        }
      }

      if (Object.keys(updateData).length === 0) {
        return NextResponse.json({ error: '수정할 항목이 없습니다' }, { status: 400 });
      }

      const { error: updateErr } = await db
        .from('offline_sales')
        .update(updateData)
        .eq('id', id);

      if (updateErr) throw updateErr;
      await recalcOutstanding(db, sale.customer_id);
      return NextResponse.json({ success: true, action: 'sale_edited' });
    }

    // --- F) 판매 재구성 (제품 추가/삭제 — 내부적으로 시리얼/재고 복원 후 재적용) ---
    if (action === 'rebuild_sale') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 수정할 수 없습니다' }, { status: 400 });
      }

      const { items: newItems, sale_info } = body as {
        items: Array<{
          product_id?: string;
          product_name: string;
          sku?: string;
          category?: string;
          quantity: number;
          unit_price: number;
          total_price: number;
          serial_ids?: string[];
          manual_serials?: string[];
        }>;
        sale_info: {
          total_amount: number;
          discount_amount?: number;
          payment_method: string;
          payment_status?: string;
          paid_amount?: number;
          sale_date?: string;
          payment_detail?: Record<string, number>;
          memo?: string;
        };
      };

      if (!newItems || newItems.length === 0) {
        return NextResponse.json({ error: '최소 1개 항목이 필요합니다' }, { status: 400 });
      }

      // 무결성 가드 — 같은 판매 안에서 동일 시리얼 중복 배정 거부 (2026-06-11)
      {
        const allManual: string[] = [];
        const allSerialIds: string[] = [];
        for (const it of newItems) {
          for (const s of (it.manual_serials || [])) { const v = String(s).trim(); if (v) allManual.push(v); }
          for (const sid of (it.serial_ids || [])) { if (sid) allSerialIds.push(sid); }
        }
        const dupManual = allManual.find((v, i) => allManual.indexOf(v) !== i);
        const dupId = allSerialIds.find((v, i) => allSerialIds.indexOf(v) !== i);
        if (dupManual || dupId) {
          return NextResponse.json({
            error: `같은 판매에 동일 시리얼이 중복 배정되었습니다${dupManual ? `: "${dupManual}"` : ''}. 품목마다 다른 시리얼을 지정해주세요.`,
            code: 'SERIAL_DUPLICATE_IN_PAYLOAD',
          }, { status: 400 });
        }
      }

      // ── STEP 1: 기존 시리얼 보존 여부 판단 ──
      // serial_ids나 manual_serials가 명시적으로 있을 때만 시리얼 재할당
      // 없으면 기존 시리얼 유지 (금액/항목만 수정하는 경우)
      const hasNewSerials = newItems.some((it) => (it.serial_ids && it.serial_ids.length > 0) || ((it as Record<string, unknown>).manual_serials as string[] || []).length > 0);

      const { data: oldSerials } = await db
        .from('product_serials')
        .select('id, product_id, warehouse_zone, previous_zone')
        .eq('offline_sale_id', id);

      if (hasNewSerials && oldSerials && oldSerials.length > 0) {
        // 새 시리얼이 명시적으로 있을 때만 기존 시리얼 복원
        for (const serial of oldSerials) {
          await db.from('product_serials').update({
            status: 'in_stock',
            warehouse_zone: serial.previous_zone || 'ready',
            previous_zone: null,
            sold_via: null,
            offline_sale_id: null,
            sold_at: null,
            sold_to_name: null,
            sold_to_phone: null,
            sale_item_id: null,
          }).eq('id', serial.id);
        }
      }

      // ── STEP 2: 기존 재고 복원 ──
      const { data: oldItems } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (oldItems) {
        const oldQtyMap: Record<string, number> = {};
        for (const item of oldItems) {
          if (item.product_id) oldQtyMap[item.product_id] = (oldQtyMap[item.product_id] || 0) + item.quantity;
        }
        for (const [productId, qty] of Object.entries(oldQtyMap)) {
          const { data: prod } = await db.from('products').select('stock_quantity, raw_stock').eq('id', productId).single();
          if (prod) {
            const hadSerials = oldSerials?.some((s: { product_id: string }) => s.product_id === productId);
            const updateData: Record<string, unknown> = { stock_quantity: (prod.stock_quantity || 0) + qty };
            if (!hadSerials) updateData.raw_stock = (prod.raw_stock || 0) + qty;
            await db.from('products').update(updateData).eq('id', productId);
          }
        }
      }

      // ── STEP 3: 기존 항목 삭제 → 새 항목 생성 (.select로 ID 확보) ──
      await db.from('offline_sale_items').delete().eq('sale_id', id);

      const saleItems = newItems.map((item) => ({
        sale_id: id,
        product_id: item.product_id || null,
        product_name: item.product_name,
        sku: item.sku || null,
        category: item.category || null,
        quantity: item.quantity,
        unit_price: item.unit_price,
        total_price: item.total_price,
      }));
      const { data: insertedSaleItems } = await db.from('offline_sale_items').insert(saleItems).select('id, product_id');

      // ── STEP 3.5: 기존 시리얼 보존 시 sale_item_id 재매칭 ──
      // 핵심: 같은 product_id 인 sale_item 이 여러 줄(qty=1 짜리 2줄 등)일 때도
      //       정확히 분배되도록 product_id 별 큐(FIFO) + quantity 만큼 pop (2026-05-17 fix)
      //       이전 버그: filter() 가 매번 같은 시리얼 배열 반환 → 마지막 sale_item 만 시리얼 보유
      if (!hasNewSerials && oldSerials && oldSerials.length > 0 && insertedSaleItems) {
        const serialQueue: Record<string, Array<{ id: string; product_id: string | null }>> = {};
        for (const sr of oldSerials as Array<{ id: string; product_id: string | null }>) {
          const key = sr.product_id || '__null__';
          if (!serialQueue[key]) serialQueue[key] = [];
          serialQueue[key].push(sr);
        }
        for (let i = 0; i < insertedSaleItems.length; i++) {
          const saleItem = insertedSaleItems[i];
          const item = newItems[i];
          if (!saleItem || !item) continue;
          const key = saleItem.product_id || '__null__';
          const queue = serialQueue[key];
          if (!queue || queue.length === 0) continue;
          const take = Math.min(item.quantity || 1, queue.length);
          for (let j = 0; j < take; j++) {
            const sr = queue.shift();
            if (sr) {
              await db.from('product_serials').update({ sale_item_id: saleItem.id }).eq('id', sr.id);
            }
          }
        }
        const leftover = Object.values(serialQueue).flat();
        if (leftover.length > 0) {
          console.warn(`[rebuild_sale] sale_id=${id} 시리얼 ${leftover.length}건 미매칭 (수량 변경 등으로 sale_item_id=null 유지)`);
        }
      }

      // ── STEP 4: 새 시리얼 할당 (아이템별 순회 + sale_item_id 매핑) ──
      if (hasNewSerials && insertedSaleItems) {
        for (let itemIdx = 0; itemIdx < newItems.length; itemIdx++) {
          const item = newItems[itemIdx];
          const saleItemId = insertedSaleItems[itemIdx]?.id || null;

          // 4-A: 기존 등록 시리얼 (serial_ids)
          const serialIds = item.serial_ids || [];
          if (serialIds.length > 0) {
            const { data: serials } = await db
              .from('product_serials')
              .select('id, warehouse_zone')
              .in('id', serialIds)
              .eq('status', 'in_stock');

            if (!serials || serials.length !== serialIds.length) {
              return NextResponse.json({ error: '일부 시리얼이 판매 불가 상태입니다' }, { status: 409 });
            }

            for (const serial of serials) {
              const { count } = await db.from('product_serials').update({
                previous_zone: serial.warehouse_zone,
                status: 'sold',
                sold_via: 'offline',
                offline_sale_id: id,
                sale_item_id: saleItemId,
                sold_at: new Date().toISOString(),
                sold_to_name: sale.customer_name,
                sold_to_phone: sale.customer_phone || null,
              }).eq('id', serial.id).eq('status', 'in_stock')
                .select('id', { count: 'exact', head: true });

              if (count === 0) {
                return NextResponse.json({ error: `시리얼 충돌 — 다시 시도해주세요` }, { status: 409 });
              }
            }
          }

          // 4-B: 직접입력/자동생성 시리얼 (manual_serials)
          const manualSerials = item.manual_serials || [];
          // Phase A 서버 안전망 — manual_serials 중 *다른 판매에 sold 상태*인 시리얼은
          // 클라이언트 동의 플래그(allow_serial_transfer) 없으면 강탈 금지 (2026-05-18)
          const allowTransfer = (body as Record<string, unknown>).allow_serial_transfer === true;
          for (const serialNumber of manualSerials) {
            if (!serialNumber.trim()) continue;
            const { data: existing } = await db
              .from('product_serials')
              .select('id, offline_sale_id, status')
              .eq('serial_number', serialNumber.trim())
              .limit(1);

            if (existing && existing.length > 0) {
              const sr = existing[0] as { id: string; offline_sale_id: string | null; status: string };
              // 다른 판매(현재 판매 id 가 아닌)에 sold 상태로 매핑되어 있으면 강탈 금지
              if (sr.offline_sale_id && sr.offline_sale_id !== id && !allowTransfer) {
                return NextResponse.json({
                  error: `시리얼 "${serialNumber.trim()}" 은 이미 다른 판매에 등록되어 있습니다. 명시적 이전 동의가 필요합니다.`,
                  code: 'SERIAL_DUPLICATE_TRANSFER_BLOCKED',
                  conflicting_serial: serialNumber.trim(),
                  conflicting_sale_id: sr.offline_sale_id,
                }, { status: 409 });
              }
              await db.from('product_serials').update({
                product_id: item.product_id || null,
                sale_item_id: saleItemId,
                status: 'sold',
                sold_via: 'offline',
                offline_sale_id: id,
                sold_at: new Date().toISOString(),
                sold_to_name: sale.customer_name,
                sold_to_phone: sale.customer_phone || null,
                previous_zone: 'ready',
              }).eq('id', sr.id);
            } else {
              await db.from('product_serials').insert({
                serial_number: serialNumber.trim(),
                product_id: item.product_id || null,
                sale_item_id: saleItemId,
                status: 'sold',
                warehouse_zone: 'ready',
                previous_zone: 'ready',
                sold_via: 'offline',
                offline_sale_id: id,
                sold_at: new Date().toISOString(),
                sold_to_name: sale.customer_name,
                sold_to_phone: sale.customer_phone || null,
              });
            }
          }
        }
      }

      // ── STEP 5: 새 재고 차감 ──
      const newQtyMap: Record<string, number> = {};
      for (const item of newItems) {
        if (item.product_id && item.quantity > 0) {
          const qty = item.serial_ids?.length || item.quantity;
          newQtyMap[item.product_id] = (newQtyMap[item.product_id] || 0) + qty;
        }
      }
      for (const [productId, qty] of Object.entries(newQtyMap)) {
        const { data: prod } = await db.from('products').select('stock_quantity, raw_stock, imweb_product_no').eq('id', productId).single();
        if (!prod) continue;
        const hasSerials = newItems.some((it) => it.product_id === productId && ((it.serial_ids && it.serial_ids.length > 0) || (it.manual_serials && it.manual_serials.length > 0)));
        const newStock = Math.max(0, (prod.stock_quantity || 0) - qty);
        const updateData: Record<string, unknown> = { stock_quantity: newStock };
        if (!hasSerials) updateData.raw_stock = Math.max(0, (prod.raw_stock || 0) - qty);
        await db.from('products').update(updateData).eq('id', productId);

        // 아임웹 동기화 — 재구성 시 차감(-)
        if (prod.imweb_product_no) {
          try {
            await updateImwebStock(Number(prod.imweb_product_no), -qty);
          } catch { /* 실패해도 계속 */ }
        }
      }

      // ── STEP 6: 판매 레코드 업데이트 ──
      const cardAmount = sale_info.payment_method === 'card' ? sale_info.total_amount
        : sale_info.payment_method === 'mixed' && sale_info.payment_detail?.card ? sale_info.payment_detail.card
        : 0;
      const supplyAmount = cardAmount > 0 ? Math.round(cardAmount / 1.1) : 0;
      const vatAmount = cardAmount > 0 ? cardAmount - supplyAmount : 0;

      await db.from('offline_sales').update({
        total_amount: sale_info.total_amount,
        discount_amount: sale_info.discount_amount || 0,
        payment_method: sale_info.payment_method,
        payment_status: sale_info.payment_status || sale.payment_status,
        paid_amount: sale_info.paid_amount ?? sale_info.total_amount,
        payment_detail: sale_info.payment_detail || null,
        sale_date: sale_info.sale_date || sale.sale_date,
        sale_channel: (sale_info as Record<string, unknown>).sale_channel || sale.sale_channel,
        memo: sale_info.memo ?? sale.memo,
        supply_amount: supplyAmount,
        vat_amount: vatAmount,
        is_vat_included: cardAmount > 0,
      }).eq('id', id);

      await recalcOutstanding(db, sale.customer_id);
      return NextResponse.json({ success: true, action: 'sale_rebuilt' });
    }

    // --- G) 070: 상담 연결 (mirror 모드 활성화) ---
    if (action === 'link_consultation') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매입니다' }, { status: 400 });
      }
      const { source_consultation_id } = body as { source_consultation_id: string };
      if (!source_consultation_id) {
        return NextResponse.json({ error: 'source_consultation_id 필수' }, { status: 400 });
      }
      // 상담 존재 + phone 일치 검증 (오링크 방지)
      const { data: consult, error: consultErr } = await db
        .from('consultations')
        .select('id, name, phone')
        .eq('id', source_consultation_id)
        .single();
      if (consultErr || !consult) {
        return NextResponse.json({ error: '상담 건을 찾을 수 없습니다' }, { status: 404 });
      }
      const salePhoneNorm = (sale.customer_phone || '').replace(/\D/g, '');
      const consultPhoneNorm = (consult.phone || '').replace(/\D/g, '');
      if (salePhoneNorm && consultPhoneNorm && salePhoneNorm !== consultPhoneNorm) {
        return NextResponse.json({ error: '판매와 상담의 고객 전화번호가 다릅니다' }, { status: 400 });
      }

      const { error: updateErr } = await db
        .from('offline_sales')
        .update({ source_consultation_id })
        .eq('id', id);
      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'consultation_linked', consultation: { id: consult.id, name: consult.name } });
    }

    // --- H) 070: 상담 연결 해제 ---
    if (action === 'unlink_consultation') {
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({ source_consultation_id: null })
        .eq('id', id);
      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'consultation_unlinked' });
    }

    return NextResponse.json({ error: '알 수 없는 action' }, { status: 400 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    console.error('[sales PATCH] error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
