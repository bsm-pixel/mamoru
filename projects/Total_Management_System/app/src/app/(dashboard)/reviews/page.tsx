'use client';

import { useState, useEffect, useCallback } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Star, Eye, EyeOff, RefreshCw, Image as ImageIcon, Award, Plus } from 'lucide-react';
import Link from 'next/link';

interface Review {
  id: string;
  review_id: string;
  created_at: string;
  type: 'consult' | 'repair' | 'purchase';
  subtype: string | null;
  name: string;
  stars: number;
  content: string;
  photo_urls: string[];
  source_id: string | null;
  status: 'pending' | 'approved' | 'hidden';
  approved_at: string | null;
  product: string | null;
  meta: Record<string, string>;
  source: string;
  is_best: boolean;
}

const TAB_FILTERS = [
  { label: '전체', value: 'all' },
  { label: '대기중', value: 'pending' },
  { label: '승인', value: 'approved' },
  { label: '숨김', value: 'hidden' },
] as const;

const TYPE_LABELS: Record<string, string> = {
  consult: '상담',
  repair: '복원수리',
  purchase: '제품구매',
};

const STATUS_STYLES: Record<string, string> = {
  pending: 'bg-yellow-50 text-yellow-700',
  approved: 'bg-green-50 text-green-700',
  hidden: 'bg-neutral-100 text-neutral-500',
};

function maskName(name: string): string {
  if (!name) return '';
  if (name.length <= 1) return name;
  if (name.length === 2) return name[0] + '*';
  return name[0] + '*'.repeat(name.length - 2) + name[name.length - 1];
}

function renderStars(n: number) {
  return Array.from({ length: 5 }, (_, i) => (
    <Star
      key={i}
      size={14}
      className={i < n ? 'fill-current text-neutral-800' : 'text-neutral-300'}
    />
  ));
}

export default function ReviewsPage() {
  const [reviews, setReviews] = useState<Review[]>([]);
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('all');
  const [lightbox, setLightbox] = useState<string | null>(null);

  const fetchReviews = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch(`/api/reviews?status=${activeTab}`);
      if (!res.ok) throw new Error('조회 실패');
      const data = await res.json();
      setReviews(data);
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  }, [activeTab]);

  useEffect(() => {
    fetchReviews();
  }, [fetchReviews]);

  const toggleStatus = async (id: string, currentStatus: string) => {
    const newStatus = currentStatus === 'approved' ? 'hidden' : 'approved';
    try {
      const res = await fetch(`/api/reviews/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: newStatus }),
      });
      if (!res.ok) throw new Error('상태 변경 실패');
      setReviews(prev =>
        prev.map(r => r.id === id ? { ...r, status: newStatus as Review['status'] } : r)
      );
    } catch (err) {
      console.error(err);
    }
  };

  const toggleBest = async (id: string, currentBest: boolean) => {
    try {
      const res = await fetch(`/api/reviews/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ is_best: !currentBest }),
      });
      if (!res.ok) throw new Error('베스트 변경 실패');
      setReviews(prev =>
        prev.map(r => r.id === id ? { ...r, is_best: !currentBest } : r)
      );
    } catch (err) {
      console.error(err);
    }
  };

  return (
    <>
      <Topbar title="리뷰관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 상단 탭 + 새로고침 */}
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div className="flex gap-1.5">
            {TAB_FILTERS.map(tab => (
              <button
                key={tab.value}
                onClick={() => setActiveTab(tab.value)}
                className={`px-3 py-1.5 rounded-full text-xs font-medium transition ${
                  activeTab === tab.value
                    ? 'bg-neutral-800 text-white'
                    : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>
          <div className="flex gap-2">
            <Link
              href="/reviews/naver"
              className="flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-medium bg-green-50 text-green-700 hover:bg-green-100 transition"
            >
              <Plus size={12} />
              네이버 리뷰 등록
            </Link>
            <button
              onClick={fetchReviews}
              disabled={loading}
              className="flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-medium bg-neutral-100 hover:bg-neutral-200 transition"
            >
              <RefreshCw size={12} className={loading ? 'animate-spin' : ''} />
              새로고침
            </button>
          </div>
        </div>

        {/* 리뷰 카드 리스트 */}
        {loading ? (
          <div className="text-center py-16 text-neutral-400 text-sm">불러오는 중...</div>
        ) : reviews.length === 0 ? (
          <div className="text-center py-16 text-neutral-400 text-sm">리뷰가 없습니다</div>
        ) : (
          <div className="grid gap-3">
            {reviews.map(review => (
              <div
                key={review.id}
                className="bg-white rounded-xl border border-neutral-100 p-4 space-y-3"
              >
                {/* 헤더: 이름 + 유형 칩 + 상태 칩 */}
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-sm">{maskName(review.name)}</span>
                    <span className="text-[10px] font-medium px-2 py-0.5 rounded-full bg-neutral-100 text-neutral-500">
                      {TYPE_LABELS[review.type] || review.type}
                    </span>
                    {review.source === 'naver' && (
                      <span className="text-[10px] font-medium px-2 py-0.5 rounded-full bg-green-50 text-green-600">
                        네이버
                      </span>
                    )}
                    <span className={`text-[10px] font-medium px-2 py-0.5 rounded-full ${STATUS_STYLES[review.status]}`}>
                      {review.status === 'pending' ? '대기' : review.status === 'approved' ? '노출중' : '숨김'}
                    </span>
                    {review.is_best && (
                      <span className="text-[10px] font-medium px-2 py-0.5 rounded-full bg-amber-50 text-amber-600">
                        BEST
                      </span>
                    )}
                  </div>
                  <span className="text-[11px] text-neutral-400">
                    {review.review_id}
                  </span>
                </div>

                {/* 별점 */}
                <div className="flex items-center gap-0.5">
                  {renderStars(review.stars)}
                </div>

                {/* 내용 */}
                <p className="text-sm text-neutral-700 leading-relaxed line-clamp-3">
                  {review.content}
                </p>

                {/* 사진 썸네일 */}
                {review.photo_urls && review.photo_urls.length > 0 && (
                  <div className="flex gap-2">
                    {review.photo_urls.map((url, i) => (
                      <button
                        key={i}
                        onClick={() => setLightbox(url)}
                        className="w-16 h-16 rounded-lg overflow-hidden border border-neutral-100 hover:opacity-80 transition"
                      >
                        <img src={url} alt="" className="w-full h-full object-cover" />
                      </button>
                    ))}
                  </div>
                )}

                {/* 하단: 날짜 + 액션 */}
                <div className="flex items-center justify-between pt-2 border-t border-neutral-50">
                  <span className="text-[11px] text-neutral-400">
                    {new Date(review.created_at).toLocaleDateString('ko-KR')}
                    {review.source_id && ` · ${review.source_id}`}
                  </span>
                  <div className="flex gap-1.5">
                    <button
                      onClick={() => toggleBest(review.id, review.is_best)}
                      className={`flex items-center gap-1 px-3 py-1 rounded-lg text-xs font-medium transition ${
                        review.is_best
                          ? 'bg-amber-50 text-amber-600 hover:bg-amber-100'
                          : 'bg-neutral-50 text-neutral-400 hover:bg-neutral-100'
                      }`}
                      title={review.is_best ? '베스트 해제' : '베스트 지정'}
                    >
                      <Award size={12} /> BEST
                    </button>
                    <button
                      onClick={() => toggleStatus(review.id, review.status)}
                      className={`flex items-center gap-1 px-3 py-1 rounded-lg text-xs font-medium transition ${
                        review.status === 'approved'
                          ? 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
                          : 'bg-green-50 text-green-700 hover:bg-green-100'
                      }`}
                    >
                      {review.status === 'approved' ? (
                        <><EyeOff size={12} /> 숨기기</>
                      ) : (
                        <><Eye size={12} /> 노출</>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* 라이트박스 */}
      {lightbox && (
        <div
          className="fixed inset-0 z-50 bg-black/70 flex items-center justify-center p-4"
          onClick={() => setLightbox(null)}
        >
          <img
            src={lightbox}
            alt="리뷰 사진"
            className="max-w-full max-h-[80vh] rounded-lg object-contain"
            onClick={e => e.stopPropagation()}
          />
        </div>
      )}
    </>
  );
}
