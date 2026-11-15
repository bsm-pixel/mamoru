import { redirect } from 'next/navigation';

// 2026-11-15 주문 대시보드(상태 5개뿐)는 주문관리와 같은 숫자를 보여 줘서 하나로 합쳤다 → 옛 주소는 주문관리로 넘긴다
export default function OrderDashboardRedirect() {
  redirect('/orders');
}
