/***** MAMORU 브랜드 소개 — GAS Web App (v1.0.0) *****/

const VERSION = 'v1.0.0';

/**
 * 웹앱 진입점 - HTML 페이지 반환
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MAMORU | 가위 전문 브랜드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}
