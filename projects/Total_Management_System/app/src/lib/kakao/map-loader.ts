/**
 * 카카오맵 SDK 동적 로더
 * 환경변수: NEXT_PUBLIC_KAKAO_MAP_KEY
 */

declare global {
  interface Window {
    kakao: {
      maps: {
        load: (callback: () => void) => void;
        LatLng: new (lat: number, lng: number) => unknown;
        Map: new (container: HTMLElement, options: { center: unknown; level: number }) => KakaoMap;
        Marker: new (options: { position: unknown; map?: KakaoMap; image?: unknown }) => KakaoMarker;
        MarkerImage: new (src: string, size: unknown, options?: unknown) => unknown;
        Size: new (w: number, h: number) => unknown;
        Point: new (x: number, y: number) => unknown;
        CustomOverlay: new (options: {
          content: string;
          position: unknown;
          yAnchor?: number;
          xAnchor?: number;
        }) => KakaoOverlay;
        event: {
          addListener: (target: unknown, type: string, handler: () => void) => void;
        };
        LatLngBounds: new () => KakaoLatLngBounds;
      };
    };
  }
}

export interface KakaoMap {
  setCenter: (latlng: unknown) => void;
  setBounds: (bounds: unknown) => void;
  setLevel: (level: number) => void;
}

export interface KakaoMarker {
  setMap: (map: KakaoMap | null) => void;
}

export interface KakaoOverlay {
  setMap: (map: KakaoMap | null) => void;
}

export interface KakaoLatLngBounds {
  extend: (latlng: unknown) => void;
}

let loaded = false;
let loading: Promise<void> | null = null;

export function loadKakaoMapSDK(): Promise<void> {
  if (loaded && window.kakao?.maps) return Promise.resolve();
  if (loading) return loading;

  const key = process.env.NEXT_PUBLIC_KAKAO_MAP_KEY;
  if (!key) {
    return Promise.reject(new Error('NEXT_PUBLIC_KAKAO_MAP_KEY 환경변수가 설정되지 않았습니다'));
  }

  loading = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = `https://dapi.kakao.com/v2/maps/sdk.js?appkey=${key}&autoload=false`;
    script.onload = () => {
      window.kakao.maps.load(() => {
        loaded = true;
        resolve();
      });
    };
    script.onerror = () => reject(new Error('카카오맵 SDK 로드 실패'));
    document.head.appendChild(script);
  });

  return loading;
}
