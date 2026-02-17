/**
 * GAS 상담 시트 → Supabase 동기화 로직
 * Google Sheets API (Service Account) → consultations 테이블 upsert
 */

import { google } from 'googleapis';
import { createServiceClient } from '@/lib/supabase/server';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';
import type { GasConsultationRow } from './types';
import { GAS_COLUMN_MAP } from './types';

const SPREADSHEET_ID = process.env.CONSULTATION_SPREADSHEET_ID!;
const SHEET_RANGE = '상담접수!A2:Z'; // 헤더 행 제외

/** Google Sheets 인증 */
function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_SERVICE_ACCOUNT_KEY?.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
}

/** GAS Status → consultation_status 매핑 */
function mapStatus(gasStatus: string): ConsultationStatus {
  const map: Record<string, ConsultationStatus> = {
    PENDING_ADMIN: 'pending_admin',
    SUGGESTED: 'suggested',
    ASSIGNED: 'assigned',
    CONFIRMED: 'confirmed',
    CANCELLED: 'cancelled',
    RESCHEDULE_REQUESTED: 'reschedule_requested',
  };
  return map[gasStatus?.toUpperCase()] || 'pending_admin';
}

/** GAS 상담방식 → consultation_type 매핑 */
function mapType(gasType: string): ConsultationType {
  if (gasType?.includes('출장')) return 'field_request';
  return 'store_visit';
}

/** 시트 행 → GasConsultationRow 변환 */
function parseRow(row: string[]): GasConsultationRow {
  const get = (key: keyof typeof GAS_COLUMN_MAP) => row[GAS_COLUMN_MAP[key]] || '';
  return {
    timestamp: get('timestamp'),
    name: get('name'),
    phone: get('phone'),
    consultType: get('consultType'),
    visitDate: get('visitDate'),
    visitTime: get('visitTime'),
    postcode: get('postcode'),
    addressRoad: get('addressRoad'),
    addressDetail: get('addressDetail'),
    addressSido: get('addressSido'),
    addressSigungu: get('addressSigungu'),
    addressRegion: get('addressRegion'),
    memo: get('memo'),
    uniqueId: get('uniqueId'),
    status: get('status'),
    dealerCode: get('dealerCode'),
    dealerName: get('dealerName'),
    suggestedDate1: get('suggestedDate1'),
    suggestedDate2: get('suggestedDate2'),
    suggestedDate3: get('suggestedDate3'),
    confirmedDate: get('confirmedDate'),
    adminNote: get('adminNote'),
    source: get('source'),
    lastUpdated: get('lastUpdated'),
    updatedBy: get('updatedBy'),
    extra: get('extra'),
  };
}

/** 전화번호 정규화 (숫자만) */
function normalizePhone(phone: string): string {
  return phone.replace(/\D/g, '');
}

/** 방문일 파싱 (다양한 포맷 대응) */
function parseVisitDate(dateStr: string): string | null {
  if (!dateStr) return null;
  // yyyy-MM-dd, yyyy/MM/dd, yyyy.MM.dd 등
  const cleaned = dateStr.replace(/[./]/g, '-').trim();
  const match = cleaned.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) {
    const [, y, m, d] = match;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return null;
}

/** 동기화 실행 */
export async function syncConsultations(): Promise<{
  success: boolean;
  synced: number;
  errors: string[];
}> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  const errors: string[] = [];
  let totalSynced = 0;

  // 동기화 로그 시작
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

  try {
    // Google Sheets에서 데이터 읽기
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_RANGE,
    });

    const rows = res.data.values || [];
    if (rows.length === 0) {
      await updateSyncLog(supabase, logEntry?.id, 'completed', 0, null);
      return { success: true, synced: 0, errors: [] };
    }

    for (const rawRow of rows) {
      try {
        const row = parseRow(rawRow);

        // uniqueId 없으면 건너뛰기
        if (!row.uniqueId) continue;

        // 딜러 매칭 (dealer_code 기준)
        let dealerId: string | null = null;
        if (row.dealerCode) {
          const { data: dealer } = await supabase
            .from('dealers')
            .select('id')
            .eq('dealer_code', row.dealerCode)
            .single();
          dealerId = dealer?.id || null;
        }

        // 고객 매칭/생성 (phone_normalized 기준)
        let customerId: string | null = null;
        if (row.phone) {
          const normalized = normalizePhone(row.phone);
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
                name: row.name,
                phone: row.phone,
                source: 'consultation',
                postcode: row.postcode || null,
                address_road: row.addressRoad || null,
                address_detail: row.addressDetail || null,
              })
              .select('id')
              .single();
            customerId = newCustomer?.id || null;
          }
        }

        // suggestions JSONB 구성
        const suggestions = [row.suggestedDate1, row.suggestedDate2, row.suggestedDate3]
          .filter(Boolean);

        // 상담 upsert
        const consultData = {
          customer_id: customerId,
          name: row.name,
          phone: row.phone,
          consultation_type: mapType(row.consultType),
          visit_date: parseVisitDate(row.visitDate),
          visit_time: row.visitTime || null,
          postcode: row.postcode || null,
          address_road: row.addressRoad || null,
          address_detail: row.addressDetail || null,
          address_sido: row.addressSido || null,
          address_sigungu: row.addressSigungu || null,
          address_region: row.addressRegion || null,
          status: mapStatus(row.status),
          memo: [row.memo, row.adminNote].filter(Boolean).join(' | ') || null,
          unique_id: row.uniqueId,
          dealer_id: dealerId,
          suggestions: suggestions.length > 0 ? { dates: suggestions } : null,
          gas_raw: row as unknown as Record<string, unknown>,
          received_at: row.timestamp ? new Date(row.timestamp).toISOString() : new Date().toISOString(),
        };

        await supabase
          .from('consultations')
          .upsert(consultData, { onConflict: 'unique_id' });

        totalSynced++;
      } catch (err) {
        const row = parseRow(rawRow);
        errors.push(`상담 ${row.uniqueId || '(ID없음)'}: ${err}`);
      }
    }

    await updateSyncLog(
      supabase,
      logEntry?.id,
      'completed',
      totalSynced,
      errors.length > 0 ? errors.join('; ') : null
    );

    return { success: true, synced: totalSynced, errors };
  } catch (err) {
    await updateSyncLog(supabase, logEntry?.id, 'failed', totalSynced, String(err));
    throw err;
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function updateSyncLog(supabase: any, logId: string | undefined, status: string, count: number, error: string | null) {
  if (!logId) return;
  await supabase
    .from('sync_log')
    .update({
      status,
      records_synced: count,
      error_message: error,
      completed_at: new Date().toISOString(),
    })
    .eq('id', logId);
}
