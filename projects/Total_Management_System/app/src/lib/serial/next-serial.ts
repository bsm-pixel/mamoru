/**
 * 다음 시리얼 번호 계산 — 서버 전용 (Supabase DB 접근).
 * 형식 M{YY}-{NNNN}, 연도내 전역 일련. 기존 8자리 숫자 시리얼과 분리(접두 M{YY}-).
 *
 * 동시성: max+1 패턴(기존 방식 유지). payload내 중복 거부 가드 + 클라이언트 reservedSerials
 *   회피가 백스톱. 연 9999개 초과 시 5자리로 자연 확장되나 string 정렬 역전 주의(현 물량 안전).
 */
import { formatSerial, parseSerial, currentYear2 } from './format';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** 현재 연도의 다음 일련번호(seq) */
export async function nextSerialSeq(db: any, year2: number = currentYear2()): Promise<number> {
  const prefix = `M${String(year2).padStart(2, '0')}-`;
  const { data } = await db
    .from('product_serials')
    .select('serial_number')
    .like('serial_number', `${prefix}%`)
    .order('serial_number', { ascending: false })
    .limit(1);
  let maxSeq = 0;
  if (data && data.length > 0) {
    const p = parseSerial(data[0].serial_number);
    if (p) maxSeq = p.seq;
  }
  return maxSeq + 1;
}

/** 다음 시리얼 문자열 1개 (M{YY}-{NNNN}) */
export async function nextSerialNumber(db: any, year2: number = currentYear2()): Promise<string> {
  return formatSerial(year2, await nextSerialSeq(db, year2));
}

/** 다음 시리얼 연속 count개 배열 */
export async function nextSerialBatch(db: any, count: number, year2: number = currentYear2()): Promise<string[]> {
  const start = await nextSerialSeq(db, year2);
  return Array.from({ length: count }, (_, i) => formatSerial(year2, start + i));
}
