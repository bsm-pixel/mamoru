/**
 * 다음 시리얼 번호 계산 — 서버 전용 (Supabase DB 접근).
 * 형식 MR{YY}{NNNNN}, **전역 누적**(연도 무관, 리셋 없음). 기존 숫자/구형식과 분리(접두 MR).
 *
 * 시작 번호 SERIAL_START(2026-06-11 사장님 지정 10816) — MR 시리얼이 하나도 없을 때 이 값부터.
 * 동시성: max+1 패턴. payload내 중복 거부 가드 + 클라이언트 reservedSerials 회피가 백스톱.
 * 99999 초과(6자리 진입) 시 string 정렬 역전 가능 — 수만 단위 미래라 그때 시퀀스 테이블로 승급.
 */
import { formatSerial, parseSerial, currentYear2 } from './format';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** 누적 시작 번호 (MR 시리얼 부재 시) */
export const SERIAL_START = 10816;

/** 전역 누적 다음 번호 (연도 무관, 리셋 없음) */
export async function nextSerialCounter(db: any): Promise<number> {
  const { data } = await db
    .from('product_serials')
    .select('serial_number')
    .like('serial_number', 'MR%')
    .order('serial_number', { ascending: false })
    .limit(5);
  let maxN = SERIAL_START - 1;
  for (const row of (data || [])) {
    const p = parseSerial(row.serial_number);
    if (p && p.seq > maxN) maxN = p.seq;
  }
  return maxN + 1;
}

/** 다음 시리얼 문자열 1개 (MR{YY}{NNNNN}) — 연도는 생성 시점 */
export async function nextSerialNumber(db: any, year2: number = currentYear2()): Promise<string> {
  return formatSerial(year2, await nextSerialCounter(db));
}

/** 다음 시리얼 연속 count개 배열 */
export async function nextSerialBatch(db: any, count: number, year2: number = currentYear2()): Promise<string[]> {
  const start = await nextSerialCounter(db);
  return Array.from({ length: count }, (_, i) => formatSerial(year2, start + i));
}
