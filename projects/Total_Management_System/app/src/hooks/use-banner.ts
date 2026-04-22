/**
 * 아임웹 배너 관리 훅 (react-query)
 * - useBanner: 단건 조회 (main_modal)
 * - useUpdateBanner: PATCH 업데이트
 * - useUploadBannerImage: 이미지 업로드
 */

import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';

export interface ImwebBanner {
  id: string;
  enabled: boolean;
  title: string | null;
  description: string | null;
  image_url: string | null;
  image_path: string | null;
  link_url: string | null;
  starts_at: string | null;
  ends_at: string | null;
  dismiss_cookie_hours: number;
  updated_at: string;
  updated_by: string | null;
}

/** 모든 배너 조회 (현재는 main_modal 1개만) */
export function useBanners() {
  return useQuery<ImwebBanner[]>({
    queryKey: ['imweb-banners'],
    queryFn: async () => {
      const res = await fetch('/api/imweb/banners');
      if (!res.ok) throw new Error('배너 조회 실패');
      const json = await res.json();
      return (json.data || []) as ImwebBanner[];
    },
  });
}

export function useUpdateBanner() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (payload: Partial<ImwebBanner> & { id: string }) => {
      const res = await fetch('/api/imweb/banners', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const json = await res.json();
      if (!res.ok || !json.ok) throw new Error(json.error || '저장 실패');
      return json.data as ImwebBanner;
    },
    onSuccess: () => {
      toast.success('배너 설정이 저장되었습니다');
      qc.invalidateQueries({ queryKey: ['imweb-banners'] });
    },
    onError: (err) => toast.error('저장 실패: ' + String(err)),
  });
}

export function useUploadBannerImage() {
  return useMutation({
    mutationFn: async ({ file, bannerId, oldPath }: { file: File; bannerId: string; oldPath?: string }) => {
      const form = new FormData();
      form.append('file', file);
      form.append('bannerId', bannerId);
      if (oldPath) form.append('oldPath', oldPath);

      const res = await fetch('/api/imweb/banners/upload', {
        method: 'POST',
        body: form,
      });
      const json = await res.json();
      if (!res.ok || !json.ok) throw new Error(json.error || '업로드 실패');
      return { url: json.url as string, path: json.path as string };
    },
    onError: (err) => toast.error('이미지 업로드 실패: ' + String(err)),
  });
}
