'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';

/**
 * 「제품」 메뉴 안의 구역 탭 (2026-11-15)
 *
 * 전엔 왼쪽 메뉴에 「제품」과 「창고·재고」가 따로 있었는데, 상세 칸이 같은 내용(가격 정보 · 빠른 작업)이라
 * 어느 쪽에서 해야 하는지 헷갈렸다 → 메뉴는 「제품」 하나, 재고 · 창고 배치도는 그 안의 탭.
 * 주소(/products · /inventory · /inventory/map)는 그대로라 예전 링크 · 즐겨찾기도 그대로 열린다.
 */
const TABS = [
  { href: '/products', label: '제품 · 가격' },
  { href: '/inventory', label: '재고' },
  { href: '/inventory/map', label: '창고 배치도' },
];

export function ProductSectionTabs() {
  const pathname = usePathname();
  const current = pathname.startsWith('/inventory/map') ? '/inventory/map' : pathname.startsWith('/inventory') ? '/inventory' : '/products';
  return (
    <nav aria-label="제품 구역" className="flex gap-1 px-4 md:px-6 border-b border-neutral-200 bg-white shrink-0">
      {TABS.map((t) => (
        <Link
          key={t.href}
          href={t.href}
          aria-current={current === t.href ? 'page' : undefined}
          className={`px-3 md:px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px whitespace-nowrap transition ${
            current === t.href ? 'border-stone-900 text-stone-900' : 'border-transparent text-neutral-500 hover:text-neutral-700'
          }`}
        >
          {t.label}
        </Link>
      ))}
    </nav>
  );
}
