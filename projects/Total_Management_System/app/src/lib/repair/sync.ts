/**
 * GAS → Supabase 복원수리 동기화 로직
 * GAS doPost(AS_CREATE) → Make/TMS webhook → Supabase upsert
 */

import { createServiceClient } from '@/lib/supabase/server';
import { calcTotalCost } from './cost-calculator';

/** GAS에서 보내는 복원수리 데이터 */
export interface RepairGasPushPayload {
  as_id: string;           // AS-YYYYMMDD-NNN
  name: string;
  phone: string;
  proceed_type?: string;   // 방문수거/카운터보관/직접전달/직접발송
  postcode?: string;
  address?: string;
  address_detail?: string;
  pickup_date?: string;
  delivery_method?: string;
  qty_mamoru?: number;
  qty_other?: number;
  memo?: string;
  service_cost?: number;
  shipping_fee?: number;
  total_amount?: number;
  received_at?: string;
  raw?: Record<string, unknown>;
}

/** 전화번호 정규화 */
function normalizePhone(phone: string): string {
  return phone.replace(/\D/g, '');
}

/** 단건 upsert */
export async function upsertRepair(payload: RepairGasPushPayload): Promise<{
  success: boolean;
  id?: string;
  error?: string;
}> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();

  try {
    if (!payload.as_id || !payload.name || !payload.phone) {
      return { success: false, error: 'as_id, name, phone 필수' };
    }

    // 고객 매칭/생성
    let customerId: string | null = null;
    const normalized = normalizePhone(payload.phone);
    const { data: existing } = await supabase
      .from('customers')
      .select('id')
      .eq('phone_normalized', normalized)
      .limit(1)
      .single();

    if (existing) {
      customerId = existing.id;
    } else {
      const { data: newCustomer } = await supabase
        .from('customers')
        .insert({
          name: payload.name,
          phone: payload.phone,
          source: 'as',
          postcode: payload.postcode || null,
          address_road: payload.address || null,
          address_detail: payload.address_detail || null,
        })
        .select('id')
        .single();
      customerId = newCustomer?.id || null;
    }

    // 비용 계산 (GAS에서 이미 계산해서 보낸 값 우선 사용)
    const qtyM = payload.qty_mamoru || 0;
    const qtyO = payload.qty_other || 0;
    const costs = calcTotalCost(qtyM, qtyO, payload.proceed_type || null);

    const repairData = {
      as_id: payload.as_id,
      customer_id: customerId,
      name: payload.name,
      phone: payload.phone,
      proceed_type: payload.proceed_type || null,
      postcode: payload.postcode || null,
      address: payload.address || null,
      address_detail: payload.address_detail || null,
      pickup_date: payload.pickup_date || null,
      delivery_method: payload.delivery_method || null,
      qty_mamoru: qtyM,
      qty_other: qtyO,
      memo: payload.memo || null,
      service_cost: payload.service_cost ?? costs.serviceCost,
      shipping_fee: payload.shipping_fee ?? costs.shippingFee,
      total_amount: payload.total_amount ?? costs.totalAmount,
      status: 'intake' as const,
      gas_raw: payload.raw || null,
      received_at: payload.received_at || new Date().toISOString(),
    };

    const { data, error } = await supabase
      .from('repairs')
      .upsert(repairData, { onConflict: 'as_id' })
      .select('id')
      .single();

    if (error) throw error;

    return { success: true, id: data.id };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

/** 배치 upsert */
export async function batchUpsertRepairs(payloads: RepairGasPushPayload[]): Promise<{
  success: boolean;
  synced: number;
  errors: string[];
}> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  let synced = 0;
  const errors: string[] = [];

  const { data: logEntry } = await supabase
    .from('sync_log')
    .insert({
      sync_type: 'repair',
      status: 'running',
      records_synced: 0,
      started_at: new Date().toISOString(),
    })
    .select()
    .single();

  for (const payload of payloads) {
    const result = await upsertRepair(payload);
    if (result.success) {
      synced++;
    } else {
      errors.push(`${payload.as_id}: ${result.error}`);
    }
  }

  if (logEntry) {
    await supabase
      .from('sync_log')
      .update({
        status: errors.length === payloads.length ? 'failed' : 'completed',
        records_synced: synced,
        error_message: errors.length > 0 ? errors.join('; ') : null,
        completed_at: new Date().toISOString(),
      })
      .eq('id', logEntry.id);
  }

  return { success: true, synced, errors };
}
