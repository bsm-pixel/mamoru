/**
 * MAMORU 간편진단 — 콘텐츠 데이터
 *
 * ★ 아임웹 page_diag.html과 동기화 포인트 ★
 * 질문/선택지/아이콘 수정 시 이 파일만 교체하면 됩니다.
 */
import type { DiagnosisQuestion, DiagnosisAnswers } from './types';

/* ─── 외부 SVG 파일 인라인화 (icons/ 폴더 → 문자열) ─── */

/** 인턴/스탭 아이콘 (Level_2.svg) */
const ICON_LEVEL2 = `<svg viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg"><defs><style>.st0{fill:#D4613E}</style></defs><g><path class="st0" d="M16.4,23.9h-1.1c-2.2,0-3.9-1.8-3.9-3.9v-1.1c0-2.2,1.8-3.9,3.9-3.9h1.1c2.2,0,3.9,1.8,3.9,3.9v1.1c0,2.2-1.8,3.9-3.9,3.9ZM15.2,15.9c-1.6,0-2.9,1.3-2.9,2.9v1.1c0,1.6,1.3,2.9,2.9,2.9h1.1c1.6,0,2.9-1.3,2.9-2.9v-1.1c0-1.6-1.3-2.9-2.9-2.9h-1.1Z"/><path class="st0" d="M14.3,19.7c-.5,0-1-.3-1.4-.8l.8-.6c.2.2.4.5.7.5s.5-.2.7-.5l.8.6c-.5.5-.9.8-1.4.8Z"/><path class="st0" d="M17.3,19.7c-.5,0-1-.3-1.4-.8l.8-.6c.2.2.4.5.7.5s.5-.2.7-.5l.8.6c-.5.5-.9.8-1.4.8Z"/><path class="st0" d="M18.5,12.8c-.3,0-.5-.2-.5-.5,0-.8.4-1.5,1.2-1.9.3-.1.4-.3.5-.5l2.3-4.6c.1-.2.4-.3.7-.2.2.1.3.4.2.7l-2.3,4.6c-.2.4-.5.7-1,.9-.4.2-.6.5-.6.9,0,.3-.2.5-.5.5,0,0,0,0,0,0Z"/><path class="st0" d="M20.1,14.9s0,0-.1,0c-.3,0-.5-.2-.5-.5s.2-.5.5-.5c.2,0,.3,0,.5,0,.4,0,.6-.2.8-.5.1-.3.2-.6,0-.8v-.3c-.3-.5-.4-1-.2-1.5l2-4.8c.1-.3.4-.4.7-.3.3.1.4.4.3.7l-2,4.8c0,.2,0,.5,0,.7v.3c.4.6.3,1.2,0,1.7-.3.5-.8.9-1.4,1-.2,0-.4,0-.6,0Z"/><path class="st0" d="M12.9,13.1s0,0-.1,0c-.3,0-.4-.4-.3-.6.1-.5,0-.9-.5-1.1-.5-.2-.8-.5-1-.9l-2.3-4.6c-.1-.2,0-.5.2-.7.2-.1.5,0,.7.2l2.3,4.6c.1.2.3.4.5.4,1,.4,1.4,1.4,1.2,2.4,0,.2-.3.4-.5.4Z"/><path class="st0" d="M11.3,14.9h0c-.2,0-.3,0-.5,0-.7-.1-1.2-.5-1.5-1-.3-.5-.3-1.2,0-1.7v-.3c.2-.2.2-.5.1-.7l-2-4.8c-.1-.3,0-.5.3-.7.3-.1.5,0,.7.3l2,4.8c.2.5.2,1,0,1.5v.3c-.3.3-.2.6,0,.8.1.3.4.4.7.5.2,0,.3,0,.4,0,.3,0,.5.2.5.5,0,.3-.2.5-.5.5Z"/><path class="st0" d="M10,20.5c-1.3,0-2.1-1.2-2.1-2.4s.2-1.2.5-1.6c.4-.5.9-.8,1.6-.8s.5.2.5.5-.2.5-.5.5-.6.1-.8.4c-.2.3-.3.6-.3,1s.2,1.4,1.1,1.4.5.2.5.5-.2.5-.5.5Z"/><path class="st0" d="M6.9,15.8c0,0-.1,0-.2,0-.7-.3-.8-1.2-.4-2.4.5-1.2,1.3-1.8,2-1.5.3.1.4.4.3.7-.1.3-.4.4-.7.3,0,0-.4.3-.7.9-.3.6-.2,1.1-.2,1.1.3.1.4.4.3.6,0,.2-.3.3-.5.3Z"/><path class="st0" d="M21.6,20.5c-.3,0-.5-.2-.5-.5s.2-.5.5-.5.6-.1.8-.4c.2-.3.3-.6.3-1s-.2-1.4-1.1-1.4-.5-.2-.5-.5.2-.5.5-.5c1.3,0,2.1,1.2,2.1,2.4s-.2,1.2-.5,1.6c-.4.5-.9.8-1.6.8Z"/><path class="st0" d="M11.7,15.2c-.2,0-.4,0-.4-.3-.3-.5-.3-1,0-1.5.2-.4.5-.8,1-1s.9-.3,1.4-.2c.5,0,1,.4,1.2.9.1.2,0,.5-.2.7-.2.1-.5,0-.7-.2-.3-.6-.9-.4-1.3-.2-.3.2-.8.6-.5,1.2.1.2,0,.5-.2.7,0,0-.2,0-.2,0Z"/><path class="st0" d="M19.6,15.4c0,0-.2,0-.3,0-.2-.1-.3-.5-.2-.7.4-.7,0-1.3-.4-1.5-.3-.2-1.1-.4-1.5.3-.1.2-.5.3-.7.2-.2-.1-.3-.5-.2-.7.7-1.1,2-1.2,2.9-.6.9.6,1.4,1.7.7,2.9,0,.1-.3.2-.4.2Z"/><path class="st0" d="M17.9,10.5c-.3,0-.5-.2-.5-.5,0-.8-.9-1.1-1.4-1.1s-1.4.2-1.4,1.1-.2.5-.5.5-.5-.2-.5-.5c0-2.7,4.8-2.7,4.8,0s-.2.5-.5.5Z"/><path class="st0" d="M24.9,15.2c-.2,0-.4-.1-.5-.4,0-.3,0-.5.3-.6,0,0,.2-.4,0-1-.2-.6-.4-.9-.5-.8-.3,0-.5,0-.6-.3,0-.3,0-.5.3-.6.4-.1.8,0,1.1.3.6.6.9,1.7.7,2.5-.1.5-.4.8-.8.9,0,0,0,0-.1,0Z"/><rect class="st0" x="15.4" y="20.8" width=".4" height="1"/><path class="st0" d="M20.3,26.4h-1c0-1.8-1.8-2.6-3.4-2.6s-3.4.8-3.4,2.6h-1c0-2.3,2.2-3.6,4.4-3.6s4.4,1.2,4.4,3.6Z"/></g><path class="st0" d="M18.2,24c-.7,0-1.5,0-2.2-.2v-1c4.5.6,7.8-.5,10.2-3.4l.8.6c-2.1,2.6-5.1,3.9-8.7,3.9Z"/><path class="st0" d="M13.8,24c-3.6,0-6.6-1.3-8.7-3.9l.8-.6c2.4,2.8,5.6,3.9,10,3.4v1c-.6.1-1.4.2-2.1.2Z"/></svg>`;

/** 디자이너 아이콘 (level_3.svg) */
const ICON_LEVEL3 = `<svg viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg"><defs><style>.st0{fill:#D4613E}</style></defs><path class="st0" d="M11,27.9h-1.3c-2.1,0-3.8-1.7-3.8-3.8v-1.3c0-2.1,1.7-3.8,3.8-3.8h1.3c2.1,0,3.8,1.7,3.8,3.8v1.3c0,2.1-1.7,3.8-3.8,3.8ZM9.6,19.9c-1.6,0-2.8,1.3-2.8,2.8v1.3c0,1.6,1.3,2.8,2.8,2.8h1.3c1.6,0,2.8-1.3,2.8-2.8v-1.3c0-1.6-1.3-2.8-2.8-2.8h-1.3Z"/><path class="st0" d="M22.1,15.2h-1.3c-2.1,0-3.8-1.7-3.8-3.8v-1.3c0-2.1,1.7-3.8,3.8-3.8h1.3c2.1,0,3.8,1.7,3.8,3.8v1.3c0,2.1-1.7,3.8-3.8,3.8ZM20.8,7.2c-1.6,0-2.8,1.3-2.8,2.8v1.3c0,1.6,1.3,2.8,2.8,2.8h1.3c1.6,0,2.8-1.3,2.8-2.8v-1.3c0-1.6-1.3-2.8-2.8-2.8h-1.3Z"/><rect class="st0" x="9.8" y="16.9" width="1" height="3"/><rect class="st0" x="7.9" y="16.3" width="1" height="3.5"/><rect class="st0" x="11.8" y="17.7" width="1" height="2.2"/><rect class="st0" x="9.9" y="9.1" width="1" height="3"/><rect class="st0" x="11.8" y="9.1" width="1" height="3.5"/><rect class="st0" x="7.9" y="9.1" width="1" height="2.2"/><path class="st0" d="M16.1,18.5s0,0,0,0c-.3,0-.6-.2-.8-.5-.1-.2,0-.5.2-.7.2-.1.5,0,.7.2,0,0,0,0,0,0,.3,0,1.7-1.1,2.8-2.5.2-.2.5-.3.7,0,.2.2.2.5,0,.7-.5.7-2.4,2.9-3.6,2.9Z"/><path class="st0" d="M19.8,17.1c-1.3,0-2.6-.5-3.6-1.8-.2-.2-.1-.5,0-.7.2-.2.5-.1.7,0,2.2,2.9,6,.5,6.2.4.2-.2.5,0,.7.1.2.2,0,.5-.1.7-.9.6-2.3,1.2-3.8,1.2Z"/><path class="st0" d="M23.4,21.2h-4c-.2,0-.4-.1-.5-.3-.5-1.5-.2-4.6.3-6.2,0-.3.4-.4.6-.3s.4.4.3.6c-.4,1.4-.6,3.9-.4,5.2h3.3c.4-1.2.3-3.9-.1-5.3,0-.3,0-.5.3-.6.3,0,.5,0,.6.3.5,1.6.8,4.9,0,6.4,0,.2-.3.3-.4.3Z"/><path class="st0" d="M20.4,23.7c-.3,0-.5-.2-.5-.5v-2.4c0-.3.2-.5.5-.5s.5.2.5.5v2.4c0,.3-.2.5-.5.5Z"/><path class="st0" d="M22.4,23.7c-.3,0-.5-.2-.5-.5v-2.3c0-.3.2-.5.5-.5s.5.2.5.5v2.3c0,.3-.2.5-.5.5Z"/><g><path class="st0" d="M15,18.9c-.3,0-.6-.1-.9-.3-.3-.2-.5-.6-.5-.9s0-.7.3-1c.2-.3.6-.5.9-.5.4,0,.7,0,1,.3.3.2.5.6.5.9s0,.7-.3,1c-.3.3-.7.5-1.1.5ZM15,16.9s0,0,0,0c-.2,0-.3,0-.4.2h0c0,.1-.1.3-.1.4s0,.3.2.4c.3.2.6.2.8,0,0-.1.1-.3.1-.4s0-.3-.2-.4c-.1,0-.2-.1-.4-.1ZM14.3,16.9h0,0Z"/><path class="st0" d="M15.4,16.2c-.3,0-.6-.1-.9-.3-.3-.2-.5-.6-.5-.9s0-.7.3-1c.5-.6,1.4-.7,2-.2.6.5.7,1.4.2,2-.3.3-.7.5-1.1.5ZM15.4,14.2c-.2,0-.4,0-.5.2,0,.1-.1.3-.1.4,0,.2,0,.3.2.4.3.2.6.2.8,0,.2-.3.2-.6,0-.8-.1,0-.2-.1-.4-.1Z"/><path class="st0" d="M16.8,19.4s0,0,0,0c-.2,0-.3-.3-.3-.5,0-.3-.4-.7-.6-.8-.2-.1-.2-.4-.1-.6.1-.2.4-.2.6-.1.1,0,1.2.8.9,1.7,0,.2-.2.3-.4.3Z"/><path class="st0" d="M14,17.2l-2-1.6.3-.7c.8,0,1.7,0,2.1-.2l.4.7c-.4.2-.9.3-1.4.4l1.1.9-.5.6Z"/><path class="st0" d="M12.5,15.7c-.2-.2-.9-.8-1.1-1-.8-.7-1.7-1.3-2.6-1.8-.2-.1,0-.4.2-.3,1.5.6,3.3,1,4.3,2.4.2.5-.4.7-.8.7h0ZM12,15c.2-.1.5-.3.7-.3,0,0-.3.6-.3.6-.7-1.2-2.3-1.9-3.6-2.5,0,0,.2-.3.2-.3,1.1.7,2.1,1.6,3,2.5h0Z"/><path class="st0" d="M12,15.1c-1.3-.1-2.6-.5-3.8-.9,0,0,.2-.3.2-.3,1.2.7,2.8,1.6,4.2,1.3,0,0-.3.6-.3.6-.1-.2-.1-.5-.1-.7h0ZM12.8,15c.2.4.5,1-.1,1.1-1.8,0-3.2-1-4.5-1.9,0,0-.1-.2,0-.3,0,0,.1,0,.2,0,.7.3,1.3.5,2,.7.8.3,1.6.3,2.5.4h0Z"/></g><path class="st0" d="M25.1,9.1s0,0,0,0c-.3,0-.4-.3-.4-.6l.4-2.3c-1,.1-1.3.5-1.4.7,0,.3-.3.5-.5.5-.3,0-.5-.2-.5-.5h0c0-.6-.1-1-.5-1.3-.6-.5-1.9-.5-2.6-.4v2c0,.3-.2.5-.5.5s-.5-.2-.5-.5v-2.4c0-.2.2-.4.4-.5.1,0,2.5-.5,3.8.6.3.2.5.5.6.8.4-.3,1.1-.5,2.2-.5s.3,0,.4.2c0,.1.1.3.1.4l-.6,3c0,.2-.3.4-.5.4Z"/><rect class="st0" x="20" y="12.5" width="1" height=".8" transform="translate(2.9 29.5) rotate(-75.6)"/><rect class="st0" x="9.8" y="24.5" width="1" height=".8" transform="translate(-16.4 28.7) rotate(-75.6)"/><path class="st0" d="M9.4,23.2c-.5-.5-.9-.5-1.4,0l-.7-.7c.9-.9,1.9-.9,2.8,0l-.7.7Z"/><path class="st0" d="M12.6,23.2c-.5-.5-.9-.5-1.4,0l-.7-.7c.4-.5.9-.7,1.4-.7h0c.5,0,1,.2,1.4.7l-.7.7Z"/><rect class="st0" x="18.5" y="10" width="1.6" height="1"/><rect class="st0" x="21.5" y="10" width="1.8" height="1"/></svg>`;

/** 틴닝 가위 아이콘 (thinning.svg) */
const ICON_THINNING = `<svg viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg"><defs><style>.cls-2{fill:none;stroke:#D4613E;stroke-miterlimit:10;stroke-width:.8px}.cls-3{fill:none;stroke:#D4613E;stroke-miterlimit:10;stroke-linecap:round;stroke-width:.5px}</style></defs><g><path class="cls-2" d="M15.13,21.75c1.08-.58,2.05-.88,2.42-1.04,3.12-1.38,4.58-1.48,6.01-.3s4.2.17,5.49-1.12c1.65-1.64,1.99-4.59-2.5-5.82-1.07-.29-2.36-.25-3.04,1.34-.84,1.96-2.85,4.08-6.73,4.81-.45.08-.86.12-1.24.13M27.85,18.71c-3.13,2.19-4.23.22-4.5-.54-.42-1.2.78-2.36,2.36-2.82,3.09-.89,3.63,2.31,2.14,3.36Z"/><path class="cls-2" d="M24.08,13.92s-.37-.63.11-.98,1.22.4,1.22.4c0,0-.84-.02-1.28.54"/><path class="cls-2" d="M25.88,2.9c.81,1.04.65,2.22.26,2.96-.12.23-.31.39-.49.5-.46.29-1.07.27-1.53-.02-2.14-1.35-4.42.72-4.42.72-1.34,1.33-1.53,3.24-1.53,3.24,0,0,1.46,2.3.28,3.33-1.18,1.04-1.95,1.23-1.95,1.23,0,0-.15,1.53-1.33,2.42-1.17.89-8.72,7.01-8.72,7.01l-3.99,2.98c-.34,5.42,12.77-5.55,12.77-5.55,0,0-.29-1.84,1.07-3.77,2.27-3.23,3.72-3.86,5.77-4.43,2.05-.57,2.91-1.73,3.65-2.99.63-1.07,1.05-2.23,1.81-4.33.63-1.73.17-2.8-.72-3.75-.6-.65-1.88-.77-.93.46ZM24.44,8.32s1.07,1.26-.86,3.01c-1.93,1.75-3.24.33-3.24.33,0,0-1.82-1.96.34-3.67,2.33-1.85,3.75.33,3.75.33Z"/><line class="cls-3" x1="3.19" y1="26.72" x2="3.66" y2="27.95"/><line class="cls-3" x1="3.94" y1="26.16" x2="4.46" y2="27.52"/><line class="cls-3" x1="4.69" y1="25.61" x2="5.25" y2="27.05"/><line class="cls-3" x1="5.47" y1="25.05" x2="6.02" y2="26.59"/><line class="cls-3" x1="6.27" y1="24.43" x2="6.8" y2="26.18"/><line class="cls-3" x1="7.01" y1="23.84" x2="7.58" y2="25.69"/><line class="cls-3" x1="7.76" y1="23.22" x2="8.42" y2="25.23"/><line class="cls-3" x1="8.51" y1="22.66" x2="9.25" y2="24.57"/><line class="cls-3" x1="9.26" y1="22.1" x2="10.02" y2="24.06"/><line class="cls-3" x1="10.04" y1="21.54" x2="10.86" y2="23.54"/><line class="cls-3" x1="10.84" y1="20.93" x2="11.57" y2="22.89"/><line class="cls-3" x1="11.58" y1="20.33" x2="12.39" y2="22.29"/></g></svg>`;

/** 슬라이싱 가위 아이콘 (Slide.svg) */
const ICON_SLIDE = `<svg viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg"><defs><style>.cls-1{fill:#D4613E;stroke:#D4613E;stroke-miterlimit:10;stroke-width:.5px}</style></defs><g><path class="cls-1" d="M21.47,15.72s-.39-.58.05-.95,1.19.31,1.19.31c0,0-.81.04-1.19.59"/><path class="cls-1" d="M23.8,15.14c-1.04-.21-2.27-.1-2.83,1.47-.68,1.92-2.58,3.15-6.24,4.08,0,0-.01,0-.02,0-.12.75-.03,1.49-.01,1.65,3.01-1.63,5.3-1.39,6.67-.39,1.45,1.05,4.02-.09,5.18-1.4,1.47-1.67,1.62-4.51-2.75-5.41ZM25.37,20.06c-2.85,2.29-4.03.47-4.33-.25-.47-1.12.6-2.3,2.09-2.83,2.9-1.04,3.61,1.99,2.25,3.08Z"/><path class="cls-1" d="M22.52,4.97c.83.95.76,2.08.43,2.81-.1.23-.27.39-.43.51-.43.31-1,.33-1.47.07-2.13-1.16-4.18.96-4.18.96-1.2,1.35-1.27,3.19-1.27,3.19,0,0,1.53,2.1.47,3.17-1.06,1.06-1.79,1.29-1.79,1.29,0,0,.17,1.82-1.12,2.39-5.76,2.56-7.93,5.2-8.62,7.51-.1.34.28.67.57.98.38.41,7.65-1.07,9.6-5.59,0,0-.23-1.44.22-2.44,1.55-3.45,2.66-3.81,4.59-4.48,1.93-.67,2.67-1.83,3.31-3.08.54-1.06.86-2.2,1.47-4.25.5-1.69,0-2.68-.91-3.55-.61-.59-1.84-.62-.86.49ZM21.47,10.24s1.1,1.14-.64,2.93c-1.74,1.79-3.07.51-3.07.51,0,0-1.85-1.77.1-3.53,2.12-1.91,3.61.09,3.61.09Z"/></g></svg>`;

/* ─── 13개 질문 데이터 ─── */

export const QUESTIONS: DiagnosisQuestion[] = [
  /* ── 1. 경력 (항상 표시) ── */
  {
    id: 'Q_STAGE',
    label: '경력',
    question: '현재 단계가\n어떻게 되시나요?',
    sub: '가장 가까운 상황을 선택해주세요',
    multiple: false,
    hasGif: false,
    condition: null,
    options: [
      { id: 'CE', label: '자격증 준비 & 취득', desc: '아직 현장 경험이 없어요', icon: '<svg viewBox="0 0 32 32" fill="#D4613E"><path d="M12.5,16.5h-.9c-2.2,0-4-1.8-4-4v-.9c0-2.2,1.8-4,4-4h.9c2.2,0,4,1.8,4,4v.9c0,2.2-1.8,4-4,4ZM11.5,8.5c-1.7,0-3,1.4-3,3v.9c0,1.7,1.4,3,3,3h.9c1.7,0,3-1.4,3-3v-.9c0-1.7-1.4-3-3-3h-.9Z"/><path d="M25.5,28.5H6.5c-2.1,0-3-5.5-3-10.7v-3.7c0-5.1,1-10.7,3-10.7h18.9c2.1,0,3,5.5,3,10.7v3.7c0,5.1-1,10.7-3,10.7ZM6.5,4.5c-.7,0-2,3.4-2,9.7v3.7c0,6.3,1.3,9.7,2,9.7h18.9c.7,0,2-3.4,2-9.7v-3.7c0-6.3-1.3-9.7-2-9.7H6.5Z"/><path d="M15.5,20c0-.1,0-3.5-3.5-3.5s-3.5,3.4-3.5,3.5h-1c0-1.6.9-4.5,4.5-4.5s4.5,2.9,4.5,4.5h-1Z"/><rect x="11.5" y="13.4" width="1" height="1.3" transform="translate(-5 21.5) rotate(-73.1)"/><path d="M24.6,21.1h-6.5c-.4,0-.6-.3-.6-.6s.3-.6.6-.6h6.5c.4,0,.6.3.6.6s-.3.6-.6.6Z"/><path d="M24.6,17.8h-6.5c-.4,0-.6-.3-.6-.6s.3-.6.6-.6h6.5c.4,0,.6.3.6.6s-.3.6-.6.6Z"/><path d="M24.6,24.5h-6.5c-.4,0-.6-.3-.6-.6s.3-.6.6-.6h6.5c.4,0,.6.3.6.6s-.3.6-.6.6Z"/><path d="M10.6,12.6c-.2,0-.4,0-.6-.2-.2-.2-.2-.4-.2-.6s0-.5.2-.6c.3-.3.9-.3,1.2,0,.2.2.2.4.2.6s0,.2,0,.3c0,.1-.1.2-.2.3-.2.2-.4.2-.6.2Z"/><path d="M13.3,12.6c-.2,0-.5,0-.6-.2,0,0-.2-.2-.2-.3,0-.1,0-.2,0-.3s0-.2,0-.3c0-.1.1-.2.2-.3,0,0,.2-.1.3-.2.2,0,.4,0,.7,0,.1,0,.2.1.3.2,0,0,.1.2.2.3,0,.1,0,.2,0,.3,0,.2,0,.4-.2.6,0,0-.2.2-.3.2-.1,0-.2,0-.3,0Z"/></svg>' },
      { id: 'IN', label: '인턴 & 스탭', desc: '현장에서 배우며 연습해요', icon: ICON_LEVEL2 },
      { id: 'DE', label: '디자이너', desc: '손님을 직접 담당해요', icon: ICON_LEVEL3 },
    ],
  },

  /* ── 2. 가위 종류 (항상 표시, 복수선택) ── */
  {
    id: 'Q_TYPE',
    label: '가위 종류',
    question: '어떤 종류의\n가위가 필요하신가요?',
    sub: '여러 개 선택 가능해요',
    multiple: true,
    hasGif: false,
    condition: null,
    options: [
      { id: 'BL', label: '블런트', desc: '사용비중이 가장 높은 메인 커트가위', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="6" cy="19" r="2.5"/><circle cx="18" cy="19" r="2.5"/><path d="M7.5 17L16 4"/><path d="M16.5 17L8 4"/></svg>' },
      { id: 'TH', label: '틴닝', desc: '양감 및 질감을 책임지는 가위', icon: ICON_THINNING },
      { id: 'LO', label: '장가위', desc: '면을 다듬는 가위의 기초 / 싱글링', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="5" cy="18" r="2"/><circle cx="5" cy="22" r="2"/><line x1="6.8" y1="17" x2="21" y2="6"/><line x1="6.8" y1="23" x2="21" y2="8"/><circle cx="10" cy="14" r="0.5" fill="#D4613E"/></svg>' },
      { id: 'SL', label: '슬라이싱', desc: '질감 테크닉', icon: ICON_SLIDE },
    ],
  },

  /* ── 3. 커트 느낌 (인턴/디자이너 + 블런트/장가위) ── */
  {
    id: 'Q_FEEL',
    label: '커트 느낌',
    question: '선호하는 커트 느낌이\n어떻게 되시나요?',
    sub: '가위를 쓸 때 원하는 느낌을 선택해주세요',
    multiple: false,
    hasGif: true,
    condition: (answers: DiagnosisAnswers) => {
      const stage = answers.Q_STAGE as string;
      const types = (answers.Q_TYPE || []) as string[];
      return (stage === 'IN' || stage === 'DE') && (types.includes('BL') || types.includes('LO'));
    },
    options: [
      { id: 'FEEL_SOFT', label: '부드러운 느낌', desc: '폭신한 커트감', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M6 19h12a4 4 0 0 0 0-8 5 5 0 0 0-9.5-2A3.5 3.5 0 0 0 6 19z"/></svg>' },
      { id: 'FEEL_POWER', label: '힘있고 강한 느낌', desc: '시원시원한 커트감', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="13,2 3,14 12,14 11,22 21,10 12,10" fill="#D4613E" stroke="#D4613E"/></svg>' },
      { id: 'FEEL_NONE', label: '아직 잘 모르겠어요', desc: '어떤것을 원하는지 감이안와요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 9c0-3.3 2.7-6 6-6s6 2.7 6 6c0 2-1 3.5-2.5 4.5-.8.5-1.5 1.2-1.5 2v.5"/><circle cx="15" cy="19" r="0.5" fill="#D4613E"/></svg>' },
    ],
  },

  /* ── 4. 커트 스타일 (디자이너 + 블런트) ── */
  {
    id: 'Q_STYLE',
    label: '커트 스타일',
    question: '주로 어떤 스타일로\n커트하시나요?',
    sub: '',
    multiple: false,
    hasGif: true,
    condition: (answers: DiagnosisAnswers) =>
      (answers.Q_STAGE as string) === 'DE' && ((answers.Q_TYPE || []) as string[]).includes('BL'),
    options: [
      { id: 'St_GO', label: '직진성있게 커트해요', desc: '닫으면서 뒤로 빼지 않아요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><line x1="2" y1="12" x2="20" y2="12"/><polyline points="16,7 22,12 16,17"/></svg>' },
      { id: 'St_BACK', label: '뒤로 빼면서 커트해요', desc: '뒤로 빼듯 조곤조곤', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M17 20V8c0-2.8-2.2-5-5-5S7 5.2 7 8"/><polyline points="3,12 7,8 11,12"/></svg>' },
      { id: 'St_NONE', label: '딱 기준이 없어요', desc: '상황에 따라 달라요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 1-9 9"/><polyline points="15,18 12,21 12,17"/><path d="M3 12a9 9 0 0 1 9-9"/><polyline points="9,6 12,3 12,7"/></svg>' },
    ],
  },

  /* ── 5. 커트 습관 (디자이너 + 블런트) ── */
  {
    id: 'Q_HABIT',
    label: '커트 습관',
    question: '분무를 하여 커트하는\nWET커트 비중이 어떻게 되세요?',
    sub: '',
    multiple: false,
    hasGif: true,
    condition: (answers: DiagnosisAnswers) =>
      (answers.Q_STAGE as string) === 'DE' && ((answers.Q_TYPE || []) as string[]).includes('BL'),
    options: [
      { id: 'HAB_WET', label: 'WET 커트 위주', desc: '항상 분무하고 커트해요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2c0 0-7 8.5-7 13a7 7 0 0 0 14 0c0-4.5-7-13-7-13z" fill="#D4613E"/></svg>' },
      { id: 'HAB_DRY', label: 'DRY 커트 위주', desc: '마른 모발 커트가 많아요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><line x1="12" y1="2" x2="12" y2="5"/><line x1="12" y1="19" x2="12" y2="22"/><line x1="2" y1="12" x2="5" y2="12"/><line x1="19" y1="12" x2="22" y2="12"/><line x1="4.9" y1="4.9" x2="7" y2="7"/><line x1="17" y1="17" x2="19.1" y2="19.1"/><line x1="4.9" y1="19.1" x2="7" y2="17"/><line x1="17" y1="7" x2="19.1" y2="4.9"/></svg>' },
      { id: 'HAB_NONE', label: '잘 모르겠어요', desc: '그때그때 다른것 같아요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="3" x2="12" y2="20"/><polygon points="8,22 16,22 12,19" fill="#D4613E"/><line x1="3" y1="8" x2="21" y2="8"/><circle cx="12" cy="3" r="1" fill="#D4613E"/><line x1="2" y1="12" x2="8" y2="12"/><path d="M2 12c0 1.5 1.3 3 3 3s3-1.5 3-3"/><line x1="2" y1="8" x2="2" y2="12"/><line x1="8" y1="8" x2="8" y2="12"/><line x1="16" y1="12" x2="22" y2="12"/><path d="M16 12c0 1.5 1.3 3 3 3s3-1.5 3-3"/><line x1="16" y1="8" x2="16" y2="12"/><line x1="22" y1="8" x2="22" y2="12"/></svg>' },
    ],
  },

  /* ── 6. 틴닝 감모량 (틴닝 선택 시) ── */
  {
    id: 'Q_TH_RATIO',
    label: '틴닝 감모량',
    question: '원하는 감모량이\n어떻게 되시나요?',
    sub: '숱이 빠지는 정도를 선택해주세요',
    multiple: false,
    hasGif: false,
    condition: (answers: DiagnosisAnswers) => ((answers.Q_TYPE || []) as string[]).includes('TH'),
    options: [
      { id: 'TH_25', label: '25% (메인 틴닝)', desc: '가장 범용적인 감모량', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="6" width="4" height="14" rx="1" fill="#D4613E"/><rect x="8" y="6" width="4" height="14" rx="1"/><rect x="14" y="6" width="4" height="14" rx="1"/><rect x="20" y="6" width="4" height="14" rx="1"/></svg>' },
      { id: 'TH_15', label: '15% (적은 감모)', desc: '질감틴닝 / 숱이 적은고객 대상으로 사용', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="6" width="4" height="14" rx="1"/><rect x="2" y="13" width="4" height="7" rx="1" fill="#D4613E"/><rect x="8" y="6" width="4" height="14" rx="1"/><rect x="14" y="6" width="4" height="14" rx="1"/><rect x="20" y="6" width="4" height="14" rx="1"/></svg>' },
      { id: 'TH_35', label: '35% (많은 감모)', desc: '빠른 양감조절', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="6" width="4" height="14" rx="1" fill="#D4613E"/><rect x="8" y="6" width="4" height="14" rx="1"/><rect x="8" y="6" width="4" height="7" rx="1" fill="#D4613E"/><rect x="14" y="6" width="4" height="14" rx="1"/><rect x="20" y="6" width="4" height="14" rx="1"/></svg>' },
    ],
  },

  /* ── 7. 틴닝 구매목적 (틴닝 선택 시) ── */
  {
    id: 'Q_TH_WHY',
    label: '틴닝 구매목적',
    question: '틴닝가위 구매 목적이\n어떻게 되시나요?',
    sub: '',
    multiple: false,
    hasGif: false,
    condition: (answers: DiagnosisAnswers) => ((answers.Q_TYPE || []) as string[]).includes('TH'),
    options: [
      { id: 'TH_WHY_NEW', label: '새로운 감모량 추가', desc: '없는 감모량을 채우고 싶어요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>' },
      { id: 'TH_WHY_SAME', label: '기존 틴닝 교체', desc: '쓰던 틴닝이 나쁘진 않았어요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 1-9 9"/><polyline points="15,18 12,21 12,17"/><path d="M3 12a9 9 0 0 1 9-9"/><polyline points="9,6 12,3 12,7"/></svg>' },
      { id: 'TH_WHY_UP', label: '업그레이드', desc: '같은 감모량 더 좋은 제품으로', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="7 10 12 5 17 10"/></svg>' },
    ],
  },

  /* ── 8. 장가위 용도 (장가위 선택 시) ── */
  {
    id: 'Q_LO_USE',
    label: '장가위 용도',
    question: '장가위 주 사용 용도가\n어떻게 될까요?',
    sub: '',
    multiple: false,
    hasGif: false,
    condition: (answers: DiagnosisAnswers) => ((answers.Q_TYPE || []) as string[]).includes('LO'),
    options: [
      { id: 'LO_BL', label: '블런트 겸용', desc: '커트가위처럼도 쓸 거예요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="6" cy="19" r="2.5"/><circle cx="18" cy="19" r="2.5"/><path d="M7.5 17L16 4"/><path d="M16.5 17L8 4"/></svg>' },
      { id: 'LO_SING', label: '싱글링 전용', desc: '싱글링 작업 전용이에요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="7" y="13" width="10" height="7" rx="2.5" ry="2.5"/><line x1="12" y1="13" x2="12" y2="3"/><path d="M8.5 13C8.5 11.5 8.5 11 8.5 10.5C8.5 10 9 9.5 9.5 10.5"/><path d="M14.5 13C14.5 11.5 14.5 11 14.5 10.5C14.5 10 15 9.5 15.5 10.5"/><path d="M7 15L5.5 13"/></svg>' },
    ],
  },

  /* ── 9. 슬라이싱 구매동기 (슬라이싱 선택 시) ── */
  {
    id: 'Q_SL_WHY',
    label: '슬라이싱 구매동기',
    question: '슬라이싱 가위 구매 동기가\n어떻게 되실까요?',
    sub: '',
    multiple: false,
    hasGif: false,
    condition: (answers: DiagnosisAnswers) => ((answers.Q_TYPE || []) as string[]).includes('SL'),
    options: [
      { id: 'SL_NEW', label: '첫 구매', desc: '슬라이싱 가위가 처음이에요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="12,3 14.5,8.5 20.5,9.2 16,13.5 17.2,19.5 12,16.5 6.8,19.5 8,13.5 3.5,9.2 9.5,8.5" fill="#D4613E" stroke="#D4613E"/></svg>' },
      { id: 'SL_SAME', label: '기존 제품 불만족', desc: '쓰던 슬라이싱이 마음에 안 들어요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="9" cy="10" r="0.8" fill="#D4613E"/><circle cx="15" cy="10" r="0.8" fill="#D4613E"/><path d="M8.5 16C9.5 14.5 14.5 14.5 15.5 16"/></svg>' },
    ],
  },

  /* ── 10. 불만족 이유 (SL_SAME 선택 시) ── */
  {
    id: 'Q_SL_SAME_WHY',
    label: '불만족 이유',
    question: '기존 슬라이싱 가위가\n어떤 점이 불만족스러우셨나요?',
    sub: '',
    multiple: false,
    hasGif: false,
    condition: (answers: DiagnosisAnswers) => (answers.Q_SL_WHY as string) === 'SL_SAME',
    options: [
      { id: 'SL_SAME_WHY_UNCOM', label: '밀리기만 하고 안 잘려요', desc: '커트 자체가 안 됨', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="17" r="2.5"/><circle cx="17" cy="17" r="2.5"/><line x1="9" y1="15.5" x2="15" y2="8"/><line x1="15" y1="15.5" x2="9" y2="8"/><line x1="8" y1="3" x2="16" y2="8" stroke-width="2"/><line x1="16" y1="3" x2="8" y2="8" stroke-width="2"/></svg>' },
      { id: 'SL_SAME_WHY_UNCOM1', label: '너무 많이 잘려나가요', desc: '모발이 과하게 잘림', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="19" r="2"/><circle cx="17" cy="19" r="2"/><line x1="8.5" y1="17.5" x2="14" y2="11"/><line x1="15.5" y1="17.5" x2="10" y2="11"/><polyline points="3 8 6.5 8"/><polyline points="1 8 4 5.5"/><polyline points="1 8 4 10.5"/><polyline points="21 8 17.5 8"/><polyline points="23 8 20 5.5"/><polyline points="23 8 20 10.5"/></svg>' },
      { id: 'SL_SAME_WHY_HAND', label: '핸들이 불편해요', desc: '사용감은 괜찮은데...', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="10" width="9" height="7" rx="2" ry="2"/><line x1="6.5" y1="10" x2="6.5" y2="6.5"/><line x1="8.5" y1="10" x2="8.5" y2="5"/><line x1="10.5" y1="10" x2="10.5" y2="5.5"/><line x1="12.5" y1="10" x2="12.5" y2="6.5"/><path d="M5 12.5L3.5 10.5"/><path d="M17 7C18 5.5 19 8.5 20 7"/><path d="M17 11C18 9.5 19 12.5 20 11"/><path d="M17 15C18 13.5 19 16.5 20 15"/></svg>' },
    ],
  },

  /* ── 11. 성별 (항상 표시) ── */
  {
    id: 'Q_GENDER',
    label: '성별',
    question: '성별은\n어떻게 되시나요?',
    sub: '손에 맞는 가위 추천을 위해 필요해요',
    multiple: false,
    hasGif: false,
    condition: null,
    options: [
      { id: 'FM', label: '여성', desc: '', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="7" r="3.5"/><path d="M8.5 5.5C8.5 5.5 8 8 7.5 11"/><path d="M15.5 5.5C15.5 5.5 16 8 16.5 11"/><path d="M7 16C7 13.5 9 12 12 12C15 12 17 13.5 17 16"/></svg>' },
      { id: 'M', label: '남성', desc: '', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="7.5" r="3.5"/><path d="M8.8 5.5C9.5 3.5 11 3 12 3C13 3 14.5 3.5 15.2 5.5"/><path d="M7 16C7 13.5 9 12.5 12 12.5C15 12.5 17 13.5 17 16"/></svg>' },
    ],
  },

  /* ── 12. 손가락 굵기 (항상 표시) ── */
  {
    id: 'Q_FING',
    label: '손가락 굵기',
    question: '가위를 사용할 손가락 굵기가\n어떤 편일까요?',
    sub: '핏에 맞는 가위 추천을 위해 필요해요',
    multiple: false,
    hasGif: false,
    condition: null,
    options: [
      { id: 'FING_NORMAL', label: '평범한 굵기', desc: '보통이에요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="14" r="3.5"/><line x1="13" y1="14" x2="13" y2="6"/><line x1="15.5" y1="14" x2="15.5" y2="5"/><line x1="18" y1="14" x2="18" y2="7"/><path d="M12 17H19V14"/></svg>' },
      { id: 'FING_THICK', label: '두꺼운 편', desc: '주변보다 두꺼워요', icon: '<svg viewBox="0 0 24 24" fill="none" stroke="#D4613E" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="6" width="14" height="12" rx="4" ry="4"/><path d="M5 12C3.5 12 3 10.5 3 9.5C3 8.5 3.5 8 5 8"/><line x1="8" y1="6" x2="8" y2="9"/><line x1="11" y1="6" x2="11" y2="9"/><line x1="14" y1="6" x2="14" y2="9"/><line x1="16.5" y1="6.5" x2="16.5" y2="9"/></svg>' },
    ],
  },
];

/* ─── 옵션 ID → 한글 라벨 매핑 ─── */

export const LABEL_MAP: Record<string, string> = {
  CE: '자격증 준비', IN: '인턴/스탭', DE: '디자이너',
  BL: '블런트', TH: '틴닝', LO: '장가위', SL: '슬라이싱',
  FEEL_SOFT: '부드러운 커트감', FEEL_POWER: '힘있고 강한 느낌', FEEL_NONE: '아직 모름',
  St_GO: '직진성 커트', St_BACK: '스트로크 커트', St_NONE: '기준 없음',
  HAB_WET: 'WET 커트', HAB_DRY: 'DRY 커트', HAB_NONE: '반반',
  TH_25: '25%', TH_15: '15%', TH_35: '35%',
  TH_WHY_NEW: '새 감모량', TH_WHY_SAME: '교체', TH_WHY_UP: '업그레이드',
  LO_BL: '블런트 겸용', LO_SING: '싱글링 전용',
  SL_NEW: '첫 구매', SL_SAME: '불만족 교체',
  SL_SAME_WHY_UNCOM: '안 잘림', SL_SAME_WHY_UNCOM1: '과다 절삭', SL_SAME_WHY_HAND: '핸들 불편',
  FM: '여성', M: '남성',
  FING_NORMAL: '보통', FING_THICK: '두꺼움',
};
