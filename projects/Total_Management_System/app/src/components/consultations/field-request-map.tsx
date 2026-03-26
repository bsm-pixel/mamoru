'use client';

import { useEffect, useRef, useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations } from '@/hooks/use-consultations';
import { loadKakaoMapSDK, type KakaoMap, type KakaoMarker, type KakaoOverlay } from '@/lib/kakao/map-loader';
import { formatPhone, CONSULTATION_STATUS_LABEL } from '@/lib/utils/format';
import type { Consultation } from '@/lib/supabase/types';

/** R2: 상태별 핀 색상 (노랑/주황/파랑) */
const PIN_CONFIG: Record<string, { color: string; icon: string }> = {
  pending_admin:        { color: '#EAB308', icon: '!' },   // 노랑 — 신규
  suggested:            { color: '#F97316', icon: '→' },   // 주황 — 제안/변경
  reschedule_requested: { color: '#F97316', icon: '↻' },   // 주황
  change_requested:     { color: '#F97316', icon: '↻' },   // 주황
  confirmed:            { color: '#3B82F6', icon: '✓' },   // 파랑 — 확정
  on_hold:              { color: '#9CA3AF', icon: '⏸' },   // 회색
};
const DEFAULT_PIN = { color: '#6B7280', icon: '?' };

/** 드롭핀 SVG — 상태별 색상 + 아이콘, dimmed 시 축소+반투명 */
function createPinSVG(status: string, dimmed = false): string {
  const { color, icon } = PIN_CONFIG[status] || DEFAULT_PIN;
  const w = dimmed ? 24 : 36;
  const h = dimmed ? 32 : 48;
  const opacity = dimmed ? 0.4 : 1;
  return `data:image/svg+xml,${encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 36 48" opacity="${opacity}">
      <filter id="s"><feDropShadow dx="0" dy="1" stdDeviation="1.5" flood-opacity="0.3"/></filter>
      <path d="M18 47C18 47 33 30 33 18A15 15 0 0 0 3 18C3 30 18 47 18 47Z" fill="${color}" stroke="white" stroke-width="2" filter="url(#s)"/>
      <circle cx="18" cy="18" r="9" fill="white" opacity="0.95"/>
      <text x="18" y="22" text-anchor="middle" font-size="14" font-weight="700" fill="${color}">${icon}</text>
    </svg>`
  )}`;
}

interface FieldRequestMapProps {
  selectedFieldId?: string | null;             // R2: 양방향 연동 — 리스트에서 선택된 ID
  onFieldSelect?: (id: string | null) => void; // R2: 지도→리스트 연동
  onSelect?: (id: string) => void;             // 마커 클릭 → 상세 패널 열기
  activeStatuses?: string[];                   // 탭 연동 — 해당 상태만 강조, 나머지 희미
}

export function FieldRequestMap({ selectedFieldId, onFieldSelect, onSelect, activeStatuses }: FieldRequestMapProps = {}) {
  const router = useRouter();
  const containerRef = useRef<HTMLDivElement>(null);
  const mapRef = useRef<KakaoMap | null>(null);
  const markersRef = useRef<KakaoMarker[]>([]);
  const overlaysRef = useRef<KakaoOverlay[]>([]);
  const [sdkReady, setSdkReady] = useState(false);
  const [sdkError, setSdkError] = useState<string | null>(null);

  // 모든 출장요청 건 (좌표 있는 건만 표시하므로 많이 불러옴)
  const { data } = useConsultations({ type: 'field_request', limit: 200 });
  const consultations = (data?.consultations || []).filter(
    (c) => c.latitude && c.longitude && c.status !== 'cancelled'
  );

  // SDK 로드
  useEffect(() => {
    loadKakaoMapSDK()
      .then(() => setSdkReady(true))
      .catch((err) => setSdkError(String(err)));
  }, []);

  // 전역 함수: 오버레이 "상세보기" 클릭 → 상세 패널 열기
  useEffect(() => {
    (window as any).__mamoruOpenDetail = (id: string) => onSelect?.(id);
    return () => { delete (window as any).__mamoruOpenDetail; };
  }, [onSelect]);

  // 지도 초기화 + 마커 렌더
  useEffect(() => {
    if (!sdkReady || !containerRef.current || consultations.length === 0) return;

    const { kakao } = window;

    // 첫 초기화
    if (!mapRef.current) {
      const center = new kakao.maps.LatLng(
        consultations[0].latitude!,
        consultations[0].longitude!
      );
      mapRef.current = new kakao.maps.Map(containerRef.current, {
        center,
        level: 10,
      });
    }

    const map = mapRef.current;

    // 기존 마커/오버레이 제거
    markersRef.current.forEach((m) => m.setMap(null));
    overlaysRef.current.forEach((o) => o.setMap(null));
    markersRef.current = [];
    overlaysRef.current = [];

    const bounds = new kakao.maps.LatLngBounds();

    consultations.forEach((c) => {
      const pos = new kakao.maps.LatLng(c.latitude!, c.longitude!);
      bounds.extend(pos);

      const dimmed = activeStatuses ? !activeStatuses.includes(c.status) : false;
      const pinSize = dimmed ? 24 : 36;
      const pinHeight = dimmed ? 32 : 48;
      const markerImage = new kakao.maps.MarkerImage(
        createPinSVG(c.status, dimmed),
        new kakao.maps.Size(pinSize, pinHeight),
        { offset: new kakao.maps.Point(pinSize / 2, pinHeight) }
      );

      const marker = new kakao.maps.Marker({
        position: pos,
        map,
        image: markerImage,
      });

      const overlay = new kakao.maps.CustomOverlay({
        content: buildOverlayHTML(c),
        position: pos,
        yAnchor: 1.4,
      });

      kakao.maps.event.addListener(marker, 'click', () => {
        // 기존 오버레이 닫기
        overlaysRef.current.forEach((o) => o.setMap(null));
        overlay.setMap(map);
        // R2: 양방향 연동 — 리스트 해당 카드 스크롤
        onFieldSelect?.(c.id);
        const el = document.getElementById(`field-${c.id}`);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      });

      markersRef.current.push(marker);
      overlaysRef.current.push(overlay);
    });

    // 모든 마커가 보이도록 범위 설정
    if (consultations.length > 1) {
      map.setBounds(bounds);
    }
  }, [sdkReady, consultations, router, activeStatuses]);

  if (sdkError) {
    return (
      <Card className="flex items-center justify-center h-80 text-sm text-neutral-400">
        {sdkError}
      </Card>
    );
  }

  if (!sdkReady) {
    return <Skeleton className="h-80 w-full rounded-xl" />;
  }

  return (
    <div className="relative h-80 md:h-[480px] lg:h-full">
      <div
        ref={containerRef}
        className="w-full h-full rounded-xl border border-neutral-200 overflow-hidden"
      />
      {/* 범례 */}
      <div className="absolute bottom-2 left-2 flex items-center gap-2 px-2.5 py-1.5 rounded-lg bg-white/90 backdrop-blur-sm shadow-sm text-[10px] text-neutral-600">
        <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-[#EAB308]" />신규</span>
        <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-[#F97316]" />제안</span>
        <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-[#3B82F6]" />확정</span>
        <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-[#9CA3AF]" />보류</span>
      </div>
    </div>
  );
}

function buildOverlayHTML(c: Consultation): string {
  const raw = (c as any).gas_raw;
  const days = (raw?.days as string)?.split(',').filter(Boolean) || [];
  const times = (raw?.timePrefs as string)?.split(',').filter(Boolean) || [];
  const prefHtml = (days.length || times.length) ? `
    <div style="display:flex;gap:3px;flex-wrap:wrap;margin-bottom:6px;">
      ${days.map(d => `<span style="padding:1px 5px;border-radius:4px;background:#EFF6FF;color:#2563EB;font-size:10px;font-weight:600;">${escapeHtml(d)}</span>`).join('')}
      ${times.map(t => `<span style="padding:1px 5px;border-radius:4px;background:#FFFBEB;color:#D97706;font-size:10px;font-weight:600;">${escapeHtml(t)}</span>`).join('')}
    </div>
  ` : '';

  return `
    <div style="background:white;border-radius:12px;padding:12px 16px;box-shadow:0 2px 12px rgba(0,0,0,.15);min-width:200px;max-width:260px;font-family:inherit;">
      <div style="font-weight:700;font-size:14px;margin-bottom:4px;">${escapeHtml(c.name)}</div>
      <div style="font-size:12px;color:#666;margin-bottom:2px;">${escapeHtml(formatPhone(c.phone))}</div>
      <div style="font-size:12px;color:#666;margin-bottom:2px;">${escapeHtml(c.address_road || '')}</div>
      <div style="font-size:11px;color:#999;margin-bottom:6px;">${CONSULTATION_STATUS_LABEL[c.status] || c.status}</div>
      ${prefHtml}
      <button onclick="window.__mamoruOpenDetail&&window.__mamoruOpenDetail('${c.id}')" style="font-size:12px;color:#C75B3F;text-decoration:none;font-weight:600;background:none;border:none;cursor:pointer;padding:0;">
        상세보기 &rarr;
      </button>
    </div>
  `;
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
