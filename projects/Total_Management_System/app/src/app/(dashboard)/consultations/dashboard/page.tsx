import { redirect } from 'next/navigation';

/** 기존 대시보드 URL 호환 — /consultations로 통합 리다이렉트 */
export default function ConsultationDashboardRedirect() {
  redirect('/consultations');
}
