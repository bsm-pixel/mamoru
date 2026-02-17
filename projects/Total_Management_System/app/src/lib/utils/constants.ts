/** 사이드바 내비게이션 항목 */
export const NAV_ITEMS = [
  { label: '대시보드', href: '/dashboard', icon: 'LayoutDashboard' },
  { label: '주문관리', href: '/orders', icon: 'ShoppingCart' },
  { label: '상담관리', href: '/consultations', icon: 'MessageSquare' },
  { label: '설정', href: '/settings', icon: 'Settings' },
] as const;

/** Phase 2+ 메뉴 (비활성) */
export const NAV_ITEMS_FUTURE = [
  { label: 'A/S', href: '/as', icon: 'Wrench' },
  { label: '제품', href: '/products', icon: 'Package' },
  { label: '고객', href: '/customers', icon: 'Users' },
] as const;

/** 택배사 코드 */
export const COURIER_CODES = [
  { code: 'LOTTE', name: '롯데택배' },
  { code: 'CJ', name: 'CJ대한통운' },
  { code: 'HANJIN', name: '한진택배' },
  { code: 'POST', name: '우체국택배' },
] as const;
