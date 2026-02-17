/**
 * GAS → Supabase 상담 동기화 로직
 * GAS 스크립트가 POST로 보낸 데이터를 Supabase에 upsert
 */

import { createServiceClient } from '@/lib/supabase/server';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

/** GAS에서 보내는 상담 데이터 타입 */
export interface GasPushPayload {
  uniqueId: string;
  name: string;
  phone: string;
  consultType?: string;
  visitDate?: string;
  visitTime?: string;
  postcode?: string;
  addressRoad?: string;
  addressDetail?: string;
  addressSido?: string;
  addressSigungu?: string;
  addressRegion?: string;
  memo?: string;
  status?: string;
  dealerCode?: string;
  dealerName?: string;
  suggestedDates?: string[];
  confirmedDate?: string;
  adminNote?: string;
  source?: string;
  receivedAt?: string;
  raw?: Record<string, unknown>;
}

/** GAS Status → consultation_status 매핑 */
function mapStatus(gasStatus?: string): ConsultationStatus {
  if (!gasStatus) return 'pending_admin';
  const map: Record<string, ConsultationStatus> = {
    PENDING_ADMIN: 'pending_admin',
    SUGGESTED: 'suggested',
    ASSIGNED: 'assigned',
    CONFIRMED: 'confirmed',
    CANCELLED: 'cancelled',
    RESCHEDULE_REQUESTED: 'reschedule_requested',
  };
  return map[gasStatus.toUpperCase()] || 'pending_admin';
}

/** GAS 상담방식 → consultation_type 매핑 */
function mapType(gasType?: string): ConsultationType {
  if (gasType?.includes('출장')) return 'field_request';
  return 'store_visit';
}

/** 방문일 파싱 */
function parseVisitDate(dateStr?: string): string | null {
  if (!dateStr) return null;
  const cleaned = dateStr.replace(/[./]/g, '-').trim();
  const match = cleaned.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) {
    const [, y, m, d] = match;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return null;
}

/** 전화번호 정규화 */
function normalizePhone(phone: string): string {
  return phone.replace(/\D/g, '');
}

/** 단건 upsert (GAS Push 수신 시 호출) */
export async function upsertConsultation(payload: GasPushPayload): Promise<{
  success: boolean;
  id?: string;
  error?: string;
}> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();

  try {
    if (!payload.uniqueId || !payload.name || !payload.phone) {
      return { success: false, error: 'uniqueId, name, phone 필수' };
    }

    // 딜러 매칭
    let dealerId: string | null = null;
    if (payload.dealerCode) {
      const { data: dealer } = await supabase
        .from('dealers')
        .select('id')
        .eq('dealer_code', payload.dealerCode)
        .single();
      dealerId = dealer?.id || null;
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
          source: 'consultation',
          postcode: payload.postcode || null,
          address_road: payload.addressRoad || null,
          address_detail: payload.addressDetail || null,
        })
        .select('id')
        .single();
      customerId = newCustomer?.id || null;
    }

    // upsert
    const consultData = {
      customer_id: customerId,
      name: payload.name,
      phone: payload.phone,
      consultation_type: mapType(payload.consultType),
      visit_date: parseVisitDate(payload.visitDate),
      visit_time: payload.visitTime || null,
      postcode: payload.postcode || null,
      address_road: payload.addressRoad || null,
      address_detail: payload.addressDetail || null,
      address_sido: payload.addressSido || null,
      address_sigungu: payload.addressSigungu || null,
      address_region: payload.addressRegion || null,
      status: mapStatus(payload.status),
      memo: [payload.memo, payload.adminNote].filter(Boolean).join(' | ') || null,
      unique_id: payload.uniqueId,
      dealer_id: dealerId,
      suggestions: payload.suggestedDates?.length ? { dates: payload.suggestedDates } : null,
      gas_raw: payload.raw || null,
      received_at: payload.receivedAt || new Date().toISOString(),
    };

    const { data, error } = await supabase
      .from('consultations')
      .upsert(consultData, { onConflict: 'unique_id' })
      .select('id')
      .single();

    if (error) throw error;

    return { success: true, id: data.id };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

/** 배치 upsert (여러 건 한번에) */
export async function batchUpsertConsultations(payloads: GasPushPayload[]): Promise<{
  success: boolean;
  synced: number;
  errors: string[];
}> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  let synced = 0;
  const errors: string[] = [];

  // sync_log 시작
  const { data: logEntry } = await supabase
    .from('sync_log')
    .insert({
      sync_type: 'consultation',
      status: 'running',
      records_synced: 0,
      started_at: new Date().toISOString(),
    })
    .select()
    .single();

  for (const payload of payloads) {
    const result = await upsertConsultation(payload);
    if (result.success) {
      synced++;
    } else {
      errors.push(`${payload.uniqueId}: ${result.error}`);
    }
  }

  // sync_log 완료
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
