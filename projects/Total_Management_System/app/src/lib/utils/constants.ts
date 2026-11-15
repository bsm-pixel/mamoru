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
      { label: '일정', href: '/schedule', icon: 'CalendarDays', matchPrefix: '/schedule' },
    ],
  },
  {
    group: '주문 · 배송',
    items: [
      { label: '주문관리', href: '/orders', icon: 'ShoppingCart', matchPrefix: '/orders' },
      { label: '반품·교환', href: '/returns', icon: 'Undo2', matchPrefix: '/returns' },
    ],
  },
  {
    group: '상담',
    items: [
      { label: '상담관리', href: '/consultations', icon: 'MessageSquare', matchPrefix: '/consultations' },
      { label: '복원수리', href: '/repairs', icon: 'Wrench', matchPrefix: '/repairs' },
    ],
  },
  {
    group: '판매',
    items: [
      { label: '판매 관리', href: '/sales', icon: 'Store', matchPrefix: '/sales' },
      // 재고판매(LS) 메뉴 숨김 — 커스텀 판매 중단, 온라인은 아임웹 아울렛으로 이관(2026-08-01).
      // 재고판매(LS) 화면·공개 접수폼은 2026-10-22 삭제(사장님 확정). EVENT 공용 백엔드·/api/stock-sale/categories 는 유지.
      { label: 'EVENT', href: '/events', icon: 'Zap', matchPrefix: '/events' },
      { label: '빠른 송장', href: '/manual-invoices', icon: 'Truck', matchPrefix: '/manual-invoices' },
      // 2026-05-26 Phase G-1: B2B거래 메뉴 삭제 — IA 통합 완료 (/sales/new?mode=b2b 진입)
      //   /deliveries 페이지 자체는 URL 직접 접근 시 작동 (B2B 깊은 시야 보존)
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
      // 2026-11-15 「창고·재고」 메뉴는 「제품」 안의 탭으로 합쳤다(components/products/product-section-tabs.tsx). 주소 /inventory 는 그대로
      { label: '매입관리', href: '/purchasing', icon: 'Truck', matchPrefix: '/purchasing' },
      { label: '샘플 소싱', href: '/sourcing', icon: 'PackageSearch', matchPrefix: '/sourcing' },
      { label: '부자재', href: '/supplies', icon: 'PackageOpen', matchPrefix: '/supplies' },
    ],
  },
  {
    group: 'CS',
    items: [
      { label: '시리얼 관리 & 조회', href: '/serials', icon: 'Search', matchPrefix: '/serials' },
      { label: '리뷰관리', href: '/reviews', icon: 'Star', matchPrefix: '/reviews' },
    ],
  },
  {
    group: '회계',
    items: [
      { label: '회계', href: '/reports', icon: 'BarChart3', matchPrefix: '/reports' },
      { label: '품목별 매출', href: '/reports/products', icon: 'Package', matchPrefix: '/reports/products' },
      { label: '경비', href: '/expenses', icon: 'Wallet', matchPrefix: '/expenses' },
      { label: '입출금', href: '/cashflow', icon: 'Building2', matchPrefix: '/cashflow' },
      { label: '세금계산서', href: '/tax-invoices', icon: 'FileSignature', matchPrefix: '/tax-invoices' },
    ],
  },
  {
    group: '',  // 시스템
    items: [
      // 2026-11-15 「달력 관리」 → 「영업시간·휴무」: 실제로는 예약 받는 시간 · 쉬는 날 설정이라 이름을 바꾸고 설정 옆으로 옮겼다(주소는 그대로)
      { label: '영업시간·휴무', href: '/consultations/calendar', icon: 'CalendarOff', matchPrefix: '/consultations/calendar' },
      { label: '설정', href: '/settings', icon: 'Settings', matchPrefix: '/settings' },
    ],
  },
];

/** 메뉴에서는 뺐지만 즐겨찾기(MY MENU)에 남아 있을 수 있는 항목 — 사이드바가 이름 · 아이콘을 찾을 때만 쓴다 */
export const NAV_LEGACY_ITEMS: NavItem[] = [
  { label: '재고', href: '/inventory', icon: 'Boxes', matchPrefix: '/inventory' },
];

/** 메뉴가 켜진 상태인지 — 「제품」은 재고 · 창고 배치도 화면에서도 켜진다 */
export function isNavActive(pathname: string, matchPrefix: string): boolean {
  if (matchPrefix === '/products') return pathname.startsWith('/products') || pathname.startsWith('/inventory');
  return pathname.startsWith(matchPrefix);
}

/** 플랫 목록 (모바일 네비 등에서 사용) */
export const NAV_ITEMS = NAV_GROUPS.flatMap((g) => g.items);

/** 택배사 코드 */
export const COURIER_CODES = [
  { code: 'LOTTE', name: '롯데택배' },
  { code: 'CJ', name: 'CJ대한통운' },
  { code: 'HANJIN', name: '한진택배' },
  { code: 'POST', name: '우체국택배' },
] as const;
