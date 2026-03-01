/** 네비게이션 항목 */
export interface NavItem {
  label: string;
  href: string;
  icon: string;
  matchPrefix: string;
}

export interface NavGroup {
  group: string;
  items: NavItem[];
}

/** 사이드바 그룹별 네비게이션 */
export const NAV_GROUPS: NavGroup[] = [
  {
    group: '',  // 홈 — 레이블 없음
    items: [
      { label: '대시보드', href: '/dashboard', icon: 'LayoutDashboard', matchPrefix: '/dashboard' },
    ],
  },
  {
    group: '영업',
    items: [
      { label: '고객', href: '/customers', icon: 'Users', matchPrefix: '/customers' },
      { label: '판매관리', href: '/sales', icon: 'Store', matchPrefix: '/sales' },
    ],
  },
  {
    group: 'CS · 수리',
    items: [
      { label: '상담관리', href: '/consultations/dashboard', icon: 'MessageSquare', matchPrefix: '/consultations' },
      { label: '복원수리', href: '/repairs/dashboard', icon: 'Wrench', matchPrefix: '/repairs' },
    ],
  },
  {
    group: '물류 · 재고',
    items: [
      { label: '주문관리', href: '/orders/dashboard', icon: 'ShoppingCart', matchPrefix: '/orders' },
      { label: '매입관리', href: '/purchasing', icon: 'Truck', matchPrefix: '/purchasing' },
      { label: '재고', href: '/inventory', icon: 'Boxes', matchPrefix: '/inventory' },
    ],
  },
  {
    group: '상품',
    items: [
      { label: '제품', href: '/products', icon: 'Package', matchPrefix: '/products' },
    ],
  },
  {
    group: '정산',
    items: [
      { label: '회계', href: '/reports', icon: 'BarChart3', matchPrefix: '/reports' },
    ],
  },
  {
    group: '',  // 시스템 — 레이블 없음
    items: [
      { label: '설정', href: '/settings', icon: 'Settings', matchPrefix: '/settings' },
    ],
  },
];

/** 플랫 목록 (모바일 네비 등에서 사용) */
export const NAV_ITEMS = NAV_GROUPS.flatMap((g) => g.items);

/** 택배사 코드 */
export const COURIER_CODES = [
  { code: 'LOTTE', name: '롯데택배' },
  { code: 'CJ', name: 'CJ대한통운' },
  { code: 'HANJIN', name: '한진택배' },
  { code: 'POST', name: '우체국택배' },
] as const;
