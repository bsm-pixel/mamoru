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

/** 사이드바 그룹별 네비게이션 — 업무 동선 순서 (IA v2) */
export const NAV_GROUPS: NavGroup[] = [
  {
    group: '',  // 홈
    items: [
      { label: '대시보드', href: '/dashboard', icon: 'LayoutDashboard', matchPrefix: '/dashboard' },
    ],
  },
  {
    group: '주문 · 배송',
    items: [
      { label: '주문관리', href: '/orders/dashboard', icon: 'ShoppingCart', matchPrefix: '/orders' },
    ],
  },
  {
    group: '상담',
    items: [
      { label: '상담관리', href: '/consultations', icon: 'MessageSquare', matchPrefix: '/consultations' },
      { label: '달력 관리', href: '/consultations/calendar', icon: 'CalendarOff', matchPrefix: '/consultations/calendar' },
      { label: '복원수리', href: '/repairs', icon: 'Wrench', matchPrefix: '/repairs' },
    ],
  },
  {
    group: '판매',
    items: [
      { label: '판매 입력', href: '/sales/new', icon: 'PackageOpen', matchPrefix: '/sales/new' },
      { label: '판매 조회', href: '/sales', icon: 'Store', matchPrefix: '/sales' },
      { label: '빠른 송장', href: '/manual-invoices', icon: 'Truck', matchPrefix: '/manual-invoices' },
      { label: 'B2B거래', href: '/deliveries', icon: 'Package', matchPrefix: '/deliveries' },
      { label: '계약서', href: '/contracts', icon: 'FileSignature', matchPrefix: '/contracts' },
    ],
  },
  {
    group: '고객',
    items: [
      { label: '고객', href: '/customers', icon: 'Users', matchPrefix: '/customers' },
      { label: 'B2B 거래처', href: '/suppliers', icon: 'Handshake', matchPrefix: '/suppliers' },
    ],
  },
  {
    group: '상품 · 재고',
    items: [
      { label: '제품', href: '/products', icon: 'Package', matchPrefix: '/products' },
      { label: '창고·재고', href: '/inventory', icon: 'Boxes', matchPrefix: '/inventory' },
      { label: '매입관리', href: '/purchasing', icon: 'Truck', matchPrefix: '/purchasing' },
      { label: '부자재', href: '/supplies', icon: 'PackageOpen', matchPrefix: '/supplies' },
    ],
  },
  {
    group: 'CS',
    items: [
      { label: '시리얼 조회', href: '/serials', icon: 'Search', matchPrefix: '/serials' },
      { label: '리뷰관리', href: '/reviews', icon: 'Star', matchPrefix: '/reviews' },
    ],
  },
  {
    group: '회계',
    items: [
      { label: '회계', href: '/reports', icon: 'BarChart3', matchPrefix: '/reports' },
      { label: '경비', href: '/expenses', icon: 'Wallet', matchPrefix: '/expenses' },
      { label: '입출금', href: '/cashflow', icon: 'Building2', matchPrefix: '/cashflow' },
      { label: '세금계산서', href: '/tax-invoices', icon: 'FileSignature', matchPrefix: '/tax-invoices' },
    ],
  },
  {
    group: '',  // 시스템
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
