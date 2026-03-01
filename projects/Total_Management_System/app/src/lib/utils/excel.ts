import * as XLSX from 'xlsx';

interface ExcelColumn {
  header: string;
  key: string;
  width?: number;
}

/**
 * 데이터 배열 → XLSX Buffer 생성
 */
export function createExcelBuffer(
  sheetName: string,
  columns: ExcelColumn[],
  rows: Record<string, unknown>[],
): Buffer {
  const headers = columns.map((c) => c.header);
  const data = rows.map((row) =>
    columns.map((c) => row[c.key] ?? '')
  );

  const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);

  // 열 너비 설정
  ws['!cols'] = columns.map((c) => ({ wch: c.width || 15 }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}
