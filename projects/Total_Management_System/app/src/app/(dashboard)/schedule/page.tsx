import { redirect } from 'next/navigation';

/**
 * 일정 — 2026-11-16 메뉴에서 뺐다(사장님 지시). 달력 · 일정 등록 · 다가오는 일정은 대시보드에 있다.
 * 즐겨찾기 · 홈 화면에 넣어 둔 옛 주소가 깨지지 않게 대시보드로 보낸다.
 */
export default function SchedulePage() {
  redirect('/dashboard');
}
