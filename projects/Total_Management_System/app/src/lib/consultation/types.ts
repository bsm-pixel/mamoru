/**
 * GAS 상담 시트 행 타입 (26개 컬럼)
 * 시트 컬럼 순서: A ~ Z
 */
export interface GasConsultationRow {
  timestamp: string;        // A: 접수시각
  name: string;             // B: 성함
  phone: string;            // C: 연락처
  consultType: string;      // D: 상담방식 (매장방문/출장요청)
  visitDate: string;        // E: 방문희망일
  visitTime: string;        // F: 방문희망시간
  postcode: string;         // G: 우편번호
  addressRoad: string;      // H: 도로명주소
  addressDetail: string;    // I: 상세주소
  addressSido: string;      // J: 시도
  addressSigungu: string;   // K: 시군구
  addressRegion: string;    // L: 지역
  memo: string;             // M: 메모/요청사항
  uniqueId: string;         // N: UniqueID
  status: string;           // O: Status
  dealerCode: string;       // P: 딜러코드
  dealerName: string;       // Q: 딜러명
  suggestedDate1: string;   // R: 제안일시1
  suggestedDate2: string;   // S: 제안일시2
  suggestedDate3: string;   // T: 제안일시3
  confirmedDate: string;    // U: 확정일시
  adminNote: string;        // V: 관리자메모
  source: string;           // W: 접수경로
  lastUpdated: string;      // X: 마지막수정
  updatedBy: string;        // Y: 수정자
  extra: string;            // Z: 비고
}

/** GAS 시트 컬럼 인덱스 (0-based) */
export const GAS_COLUMN_MAP = {
  timestamp: 0,
  name: 1,
  phone: 2,
  consultType: 3,
  visitDate: 4,
  visitTime: 5,
  postcode: 6,
  addressRoad: 7,
  addressDetail: 8,
  addressSido: 9,
  addressSigungu: 10,
  addressRegion: 11,
  memo: 12,
  uniqueId: 13,
  status: 14,
  dealerCode: 15,
  dealerName: 16,
  suggestedDate1: 17,
  suggestedDate2: 18,
  suggestedDate3: 19,
  confirmedDate: 20,
  adminNote: 21,
  source: 22,
  lastUpdated: 23,
  updatedBy: 24,
  extra: 25,
} as const;
