'use client';

import dynamic from 'next/dynamic';

/**
 * 처음엔 안 보이는 큰 창 — 열 때 받는다 (2026-11-21)
 *
 * 라벨 · 꼬리표 출력(바코드 그리기 포함) · 준비표 · 발주서 인쇄 창은 버튼을 눌러야 보이는데,
 * 목록 · 상세 화면에 함께 묶여 있어 화면을 열 때마다(그리고 메뉴가 다른 화면을 미리 받을 때마다) 같이 내려왔다.
 * 여기서 가져다 쓰면 창을 처음 여는 순간에만 받는다. 이름과 쓰는 법(props)은 원래 것과 같다.
 * 새로 큰 창을 만들면 여기에 한 줄 더하고, 화면에서는 이 파일에서 가져온다.
 */
export const LabelPrintModal = dynamic(() => import('@/components/labels/label-print-modal').then((m) => m.LabelPrintModal), { ssr: false });
export const RepairTagPrintModal = dynamic(() => import('@/components/labels/repair-tag-print-modal').then((m) => m.RepairTagPrintModal), { ssr: false });
export const PrepSheetModal = dynamic(() => import('@/components/sales/prep-sheet-modal').then((m) => m.PrepSheetModal), { ssr: false });
export const RepairPrepSheetModal = dynamic(() => import('@/components/repairs/repair-prep-sheet-modal').then((m) => m.RepairPrepSheetModal), { ssr: false });
export const UnifiedPrepModal = dynamic(() => import('@/components/shared/unified-prep-modal').then((m) => m.UnifiedPrepModal), { ssr: false });
export const POPrintModal = dynamic(() => import('@/components/purchasing/po-print-modal').then((m) => m.POPrintModal), { ssr: false });
