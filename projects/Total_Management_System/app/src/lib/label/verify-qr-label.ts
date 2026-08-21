import { mmToDots, fQR, fText, buildLabel } from './zpl';

/**
 * 20×20mm 정품인증 QR 라벨 (실버 라벨지 부착용).
 *   상단 = 정사각 QR (verify URL), 하단 = 모델명(영문·숫자 SKU).
 *   QR·텍스트 모두 ZPL 네이티브(^BQN / ^A0N) → 라이브러리 불필요, 선명.
 *
 * ⚠️ mag(QR 배율)·위치는 URL 길이·DPI에 따라 달라져 근사값이다.
 *    실물 1장 출력해 스캔·가독성 확인 후 미세조정할 것. (ZT231 = 203dpi 기본)
 */
export function buildVerifyQrZpl(verifyUrl: string, model: string, dpi = 203, copies = 1): string {
  const sizeMm = 20;
  const dots = mmToDots(sizeMm, dpi); // 203dpi → 160

  // QR: 55자 내외 URL ≈ 버전3~4(29~33 모듈). 203dpi는 mag 3(≈12mm), 300dpi는 mag 4.
  const mag = dpi >= 300 ? 4 : 3;
  const qrModulesApprox = 33;
  const qrPx = qrModulesApprox * mag;
  const qrX = Math.max(0, Math.round((dots - qrPx) / 2));
  const qrY = mmToDots(1.2, dpi);

  // 모델명: 하단 중앙 근사. 글자높이 ≈ 2.6mm
  const txtH = mmToDots(2.6, dpi);
  const txtW = Math.round(txtH * 0.58);
  const txtY = dots - txtH - mmToDots(1.4, dpi);
  const cleanModel = String(model || '').slice(0, 14); // 20mm 폭 한계
  const estW = cleanModel.length * txtW;
  const txtX = Math.max(2, Math.round((dots - estW) / 2));

  const fields = [
    fQR(qrX, qrY, verifyUrl, mag),
    cleanModel ? fText(txtX, txtY, cleanModel, txtH, txtW) : '',
  ];
  return buildLabel({ widthMm: sizeMm, heightMm: sizeMm, dpi }, fields, copies);
}
