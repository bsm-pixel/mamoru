'use client';

import { useEffect, useRef, useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations } from '@/hooks/use-consultations';
import { loadKakaoMapSDK, type KakaoMap, type KakaoMarker, type KakaoOverlay } from '@/lib/kakao/map-loader';
import { formatPhone, CONSULTATION_STATUS_LABEL } from '@/lib/utils/format';
import type { Consultation } from '@/lib/supabase/types';

/** 상태별 핀 색상 (SVG circle fill) */
const PIN_COLORS: Record<string, string> = {
  pending_admin: '#FBBF24', // 노란색
  suggested: '#FB923C',     // 오렌지
  reschedule_requested: '#FBBF24',
  confirmed: '#3B82F6',     // 파란색
  on_hold: '#9CA3AF',       // 회색
};

function createPinSVG(color: string, pulse = false): string {
  const animation = pulse
    ? '<animate attributeName="r" values="8;12;8" dur="1.5s" repeatCount="indefinite"/>'
    : '';
  return `data:image/svg+xml,${encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 28 28">
      <circle cx="14" cy="14" r="10" fill="${color}" stroke="white" stroke-width="2">${animation}</circle>
    </svg>`
  )}`;
}

export function FieldRequestMap() {
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

      const color = PIN_COLORS[c.status] || '#9CA3AF';
      const pulse = c.status === 'pending_admin';
      const markerImage = new kakao.maps.MarkerImage(
        createPinSVG(color, pulse),
        new kakao.maps.Size(28, 28)
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
      });

      markersRef.current.push(marker);
      overlaysRef.current.push(overlay);
    });

    // 모든 마커가 보이도록 범위 설정
    if (consultations.length > 1) {
      map.setBounds(bounds);
    }
  }, [sdkReady, consultations, router]);

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
    <div
      ref={containerRef}
      className="w-full h-80 md:h-[480px] rounded-xl border border-neutral-200 overflow-hidden"
    />
  );
}

function buildOverlayHTML(c: Consultation): string {
  return `
    <div style="background:white;border-radius:12px;padding:12px 16px;box-shadow:0 2px 12px rgba(0,0,0,.15);min-width:200px;font-family:inherit;">
      <div style="font-weight:700;font-size:14px;margin-bottom:4px;">${escapeHtml(c.name)}</div>
      <div style="font-size:12px;color:#666;margin-bottom:2px;">${escapeHtml(formatPhone(c.phone))}</div>
      <div style="font-size:12px;color:#666;margin-bottom:2px;">${escapeHtml(c.address_road || '')}</div>
      <div style="font-size:11px;color:#999;margin-bottom:8px;">${CONSULTATION_STATUS_LABEL[c.status] || c.status}</div>
      <a href="/consultations/${c.id}" style="font-size:12px;color:#C75B3F;text-decoration:none;font-weight:600;">
        상세보기 &rarr;
      </a>
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
