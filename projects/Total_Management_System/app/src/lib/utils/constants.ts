/** 사이드바 내비게이션 항목 */
export const NAV_ITEMS = [
  { label: '대시보드', href: '/dashboard', icon: 'LayoutDashboard', matchPrefix: '/dashboard' },
  { label: '주문관리', href: '/orders/dashboard', icon: 'ShoppingCart', matchPrefix: '/orders' },
  { label: '상담관리', href: '/consultations/dashboard', icon: 'MessageSquare', matchPrefix: '/consultations' },
  { label: '복원수리', href: '/repairs/dashboard', icon: 'Wrench', matchPrefix: '/repairs' },
  { label: '판매관리', href: '/sales', icon: 'Store', matchPrefix: '/sales' },  // R5: 오프라인 판매
  { label: '계약서', href: '/contracts', icon: 'FileSignature', matchPrefix: '/contracts' },  // R6: 전자 계약서
  { label: '제품', href: '/products', icon: 'Package', matchPrefix: '/products' },  // R7: 시리얼넘버/바코드
  { label: '설정', href: '/settings', icon: 'Settings', matchPrefix: '/settings' },
] as const;

/** 향후 메뉴 (비활성) */
export const NAV_ITEMS_FUTURE = [
  { label: '고객', href: '/customers', icon: 'Users' },
] as const;

/** 택배사 코드 */
export const COURIER_CODES = [
  { code: 'LOTTE', name: '롯데택배' },
  { code: 'CJ', name: 'CJ대한통운' },
  { code: 'HANJIN', name: '한진택배' },
  { code: 'POST', name: '우체국택배' },
] as const;
