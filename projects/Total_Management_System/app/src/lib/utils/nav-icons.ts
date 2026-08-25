import {
  LayoutDashboard, ShoppingCart, Settings, MessageSquare, Wrench, Store,
  Package, Users, Truck, Boxes, BarChart3, Star, Building2, Handshake,
  PackageOpen, PackageSearch, Search, FileSignature, Wallet, CalendarOff,
  CalendarDays, Zap, Tag, Undo2,
  type LucideIcon,
} from 'lucide-react';

/**
 * 네비게이션 아이콘 문자열 → lucide 컴포넌트 매핑 (SSOT)
 * sidebar.tsx · mobile-nav.tsx 공용 — 두 곳에 중복 정의하던 것을 단일화.
 * constants.ts NAV_GROUPS/NAV_ITEMS 의 `icon` 문자열과 키가 일치해야 한다.
 */
export const NAV_ICON_MAP: Record<string, LucideIcon> = {
  LayoutDashboard, ShoppingCart, Settings, MessageSquare, Wrench, Store,
  Package, Users, Truck, Boxes, BarChart3, Star, Building2, Handshake,
  PackageOpen, PackageSearch, Search, FileSignature, Wallet, CalendarOff,
  CalendarDays, Zap, Tag, Undo2,
};
