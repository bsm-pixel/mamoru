'use client';

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { Camera, Trash2, Plus, X } from 'lucide-react';
import toast from 'react-hot-toast';
import { resizeImage } from '@/lib/utils/resize-image';
import { CameraCapture } from '@/components/ui/camera-capture';

interface Photo {
  id: string;
  file_name: string;
  file_path: string;
  memo: string | null;
  url: string;
  created_at: string;
}

export function RepairPhotos({ repairId }: { repairId: string }) {
  const queryClient = useQueryClient();
  const [preview, setPreview] = useState<string | null>(null);
  const [cameraOpen, setCameraOpen] = useState(false);

  const { data, isLoading } = useQuery({
    queryKey: ['repair-photos', repairId],
    queryFn: async () => {
      const res = await fetch(`/api/repair/photos?repair_id=${repairId}`);
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      return json.photos as Photo[];
    },
    staleTime: 60_000,
  });

  const upload = useMutation({
    mutationFn: async (file: File) => {
      const form = new FormData();
      form.append('repair_id', repairId);
      form.append('file', file);
      const res = await fetch('/api/repair/photos', { method: 'POST', body: form });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('사진 업로드 완료');
      queryClient.invalidateQueries({ queryKey: ['repair-photos', repairId] });
    },
    onError: (err) => toast.error('업로드 실패: ' + String(err)),
  });

  const remove = useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/repair/photos?id=${id}`, { method: 'DELETE' });
      if (!res.ok) throw new Error(await res.text());
    },
    onSuccess: () => {
      toast.success('삭제 완료');
      queryClient.invalidateQueries({ queryKey: ['repair-photos', repairId] });
    },
    onError: (err) => toast.error('삭제 실패: ' + String(err)),
  });

  const photos = data || [];

  async function handleCapture(file: File) {
    const resized = await resizeImage(file, 1200, 0.8);
    upload.mutate(resized);
  }

  return (
    <Card>
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2">
          <Camera size={16} className="text-neutral-400" />
          <span className="text-xs font-semibold text-neutral-500">사진 ({photos.length})</span>
        </div>
        <Button
          size="sm"
          variant="secondary"
          onClick={() => setCameraOpen(true)}
          disabled={upload.isPending}
        >
          <Plus size={14} />
          {upload.isPending ? '업로드 중...' : '사진 추가'}
        </Button>
      </div>

      <CameraCapture open={cameraOpen} onClose={() => setCameraOpen(false)} onCapture={handleCapture} />

      {isLoading ? (
        <div className="grid grid-cols-3 gap-2">
          {Array.from({ length: 3 }).map((_, i) => (
            <Skeleton key={i} className="aspect-square rounded-lg" />
          ))}
        </div>
      ) : photos.length === 0 ? (
        <div className="flex flex-col items-center py-6 text-neutral-300">
          <Camera size={24} className="mb-1 opacity-40" />
          <p className="text-xs">등록된 사진이 없습니다</p>
        </div>
      ) : (
        <div className="grid grid-cols-3 gap-2">
          {photos.map((p) => (
            <div key={p.id} className="relative group">
              <img
                src={p.url}
                alt={p.file_name}
                className="aspect-square rounded-lg object-cover cursor-pointer hover:opacity-90 transition"
                onClick={() => setPreview(p.url)}
              />
              <button
                onClick={() => remove.mutate(p.id)}
                className="absolute top-1 right-1 w-7 h-7 rounded-full bg-black/50 text-white flex items-center justify-center opacity-100 md:opacity-0 md:group-hover:opacity-100 transition"
              >
                <Trash2 size={13} />
              </button>
            </div>
          ))}
        </div>
      )}

      {/* 풀스크린 프리뷰 */}
      {preview && (
        <div
          className="fixed inset-0 z-50 bg-black/90 flex items-center justify-center"
          onClick={() => setPreview(null)}
        >
          <button
            className="absolute top-4 right-4 w-10 h-10 rounded-full bg-white/20 flex items-center justify-center"
            onClick={() => setPreview(null)}
          >
            <X size={20} className="text-white" />
          </button>
          <img src={preview} alt="" className="max-w-full max-h-full object-contain" />
        </div>
      )}
    </Card>
  );
}
