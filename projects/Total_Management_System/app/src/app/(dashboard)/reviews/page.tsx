'use client';

import { useState, useEffect, useCallback, useRef } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Star, Eye, EyeOff, RefreshCw, Image as ImageIcon, Award, Plus, Pencil, X, Trash2 } from 'lucide-react';
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
  { label: '약속 대기', value: 'promised' },
] as const;

const TYPE_LABELS: Record<string, string> = {
  consult: '상담',
  repair: '복원수리',
  purchase: '제품구매',
};

const SUBTYPE_LABELS: Record<string, string> = {
  store_visit: '직접방문', field_request: '출장', talk_consult: '톡상담',
  restoration: '복원수리', direct_visit: '직접방문', pickup: '방문수거',
  parcel_pickup: '방문수거', self_ship: '직접발송',
  store: '매장', field: '출장', offline: '오프라인', online: '온라인', talk: '톡상담', // sale_channel 값(2026-07-17 4분류) + 레거시 호환
};

function getChipLabel(review: Review): string {
  const type = TYPE_LABELS[review.type] || review.type;
  const sub = review.subtype ? (SUBTYPE_LABELS[review.subtype] || review.subtype) : '';
  if (review.type === 'consult' && sub) return `${type}·${sub}`;
  return type;
}

function getSubChip(review: Review): string | null {
  if (review.type !== 'repair') return null;
  if (!review.subtype || review.subtype === 'restoration') return null; // 기본값은 type과 중복이므로 생략
  return SUBTYPE_LABELS[review.subtype] || review.subtype;
}

const STATUS_STYLES: Record<string, string> = {
  pending: 'bg-yellow-50 text-yellow-700',
  approved: 'bg-green-50 text-green-700',
  hidden: 'bg-neutral-100 text-neutral-500',
};

function renderStars(n: number) {
  return Array.from({ length: 5 }, (_, i) => (
    <Star
      key={i}
      size={14}
      className={i < n ? 'fill-current text-neutral-800' : 'text-neutral-300'}
    />
  ));
}

const TYPE_OPTIONS: { value: string; label: string }[] = [
  { value: 'consult', label: '컨설팅상담' },
  { value: 'repair', label: '복원수리' },
  { value: 'purchase', label: '제품구매' },
];

interface PromisedItem {
  source: 'consultation' | 'repair' | 'sale';
  id: string;
  displayId: string;
  customerName: string;
  customerPhone: string | null;
  typeLabel: string;
  promisedAt: string;
  requestSentAt: string | null;
}

interface RelatedActivity {
  source: 'consultation' | 'repair' | 'sale';
  id: string;
  displayId: string;
  typeLabel: string;
  promisedAt: string | null;
  requestSentAt: string | null;
  submittedAt: string | null;
}

export default function ReviewsPage() {
  const [reviews, setReviews] = useState<Review[]>([]);
  const [promisedItems, setPromisedItems] = useState<PromisedItem[]>([]);
  const [relatedByItem, setRelatedByItem] = useState<Record<string, RelatedActivity[]>>({});
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('all');
  const [lightbox, setLightbox] = useState<string | null>(null);
  const [editReview, setEditReview] = useState<Review | null>(null);
  const [editForm, setEditForm] = useState<Partial<Review>>({});
  const [saving, setSaving] = useState(false);
  const [photoUploading, setPhotoUploading] = useState(false);
  const [autoApprove, setAutoApprove] = useState(false);
  const [autoLoaded, setAutoLoaded] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<string | null>(null);
  const editPhotoRef = useRef<HTMLInputElement>(null);

  // 자동 노출 토글 로드
  useEffect(() => {
    fetch('/api/settings').then(r => r.json()).then(d => {
      // API가 { key: value } 맵을 직접 반환
      setAutoApprove(d['review.auto_approve'] === 'true');
      setAutoLoaded(true);
    }).catch(() => setAutoLoaded(true));
  }, []);

  const toggleAutoApprove = async () => {
    const newVal = !autoApprove;
    setAutoApprove(newVal);
    await fetch('/api/settings', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ items: [{ key: 'review.auto_approve', value: String(newVal) }] }),
    });
  };

  const handleDelete = async (id: string) => {
    try {
      const res = await fetch(`/api/reviews/${id}`, { method: 'DELETE' });
      if (!res.ok) throw new Error();
      setReviews(prev => prev.filter(r => r.id !== id));
      setDeleteTarget(null);
    } catch { alert('삭제 실패'); }
  };

  const fetchReviews = useCallback(async () => {
    setLoading(true);
    try {
      if (activeTab === 'promised') {
        const res = await fetch('/api/reviews/promised');
        if (!res.ok) throw new Error('약속 대기 조회 실패');
        const data = await res.json();
        setPromisedItems(data.items || []);
      } else {
        const res = await fetch(`/api/reviews?status=${activeTab}`);
        if (!res.ok) throw new Error('조회 실패');
        const data = await res.json();
        setReviews(data);
      }
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  }, [activeTab]);

  useEffect(() => {
    fetchReviews();
  }, [fetchReviews]);

  // 약속 대기 탭: 각 row의 같은 phone 다른 source 활동 fetch (정보 표시용)
  useEffect(() => {
    if (activeTab !== 'promised' || promisedItems.length === 0) {
      setRelatedByItem({});
      return;
    }
    let cancelled = false;
    (async () => {
      const results = await Promise.all(
        promisedItems.map(async (it) => {
          const key = `${it.source}-${it.id}`;
          if (!it.customerPhone) return [key, [] as RelatedActivity[]] as const;
          try {
            const params = new URLSearchParams({
              phone: it.customerPhone,
              excludeSource: it.source,
              excludeId: it.id,
            });
            const res = await fetch(`/api/reviews/related-activity?${params}`);
            if (!res.ok) return [key, [] as RelatedActivity[]] as const;
            const data = await res.json();
            return [key, (data.items || []) as RelatedActivity[]] as const;
          } catch {
            return [key, [] as RelatedActivity[]] as const;
          }
        })
      );
      if (cancelled) return;
      const map: Record<string, RelatedActivity[]> = {};
      for (const [k, v] of results) map[k] = v;
      setRelatedByItem(map);
    })();
    return () => { cancelled = true; };
  }, [activeTab, promisedItems]);

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

  /* ── 수정 모달 열기 ── */
  const openEdit = (review: Review) => {
    setEditReview(review);
    setEditForm({
      name: review.name,
      type: review.type,
      stars: review.stars,
      content: review.content,
      photo_urls: [...(review.photo_urls || [])],
      product: review.product || '',
      created_at: review.created_at?.slice(0, 10) || '',
      is_best: review.is_best,
    });
  };

  /* ── 수정 사진 업로드 ── */
  const handleEditPhotoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    setPhotoUploading(true);
    try {
      const { resizeImage } = await import('@/lib/utils/resize-image');
      const resized = await Promise.all(Array.from(files).map(f => resizeImage(f, 1200, 0.8)));
      const formData = new FormData();
      resized.forEach(f => formData.append('files', f));
      const res = await fetch('/api/reviews/upload-bulk', { method: 'POST', body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      const urls = (data.results || []).filter((r: { url?: string }) => r.url).map((r: { url: string }) => r.url);
      setEditForm(prev => ({ ...prev, photo_urls: [...(prev.photo_urls || []), ...urls] }));
    } catch (err) {
      alert('사진 업로드 실패: ' + String(err));
    } finally {
      setPhotoUploading(false);
      e.target.value = '';
    }
  };

  /* ── 수정 저장 ── */
  const saveEdit = async () => {
    if (!editReview) return;
    setSaving(true);
    try {
      const res = await fetch(`/api/reviews/${editReview.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: editForm.name,
          type: editForm.type,
          stars: editForm.stars,
          content: editForm.content,
          photo_urls: editForm.photo_urls,
          product: editForm.product || null,
          created_at: editForm.created_at || editReview.created_at,
          is_best: editForm.is_best,
        }),
      });
      if (!res.ok) throw new Error('수정 실패');
      const updated = await res.json();
      setReviews(prev => prev.map(r => r.id === editReview.id ? { ...r, ...updated } : r));
      setEditReview(null);
    } catch (err) {
      alert('수정 실패: ' + String(err));
    } finally {
      setSaving(false);
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
          <div className="flex items-center gap-2">
            {/* 자동 노출 토글 */}
            {autoLoaded && (
              <button
                onClick={toggleAutoApprove}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition ${
                  autoApprove ? 'bg-green-50 text-green-700' : 'bg-neutral-100 text-neutral-500'
                }`}
              >
                <span className={`w-2 h-2 rounded-full ${autoApprove ? 'bg-green-500' : 'bg-neutral-300'}`} />
                즉시 노출 {autoApprove ? 'ON' : 'OFF'}
              </button>
            )}
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

        {/* 약속 대기 탭 — 별도 리스트 */}
        {activeTab === 'promised' && (
          loading ? (
            <div className="text-center py-16 text-neutral-400 text-sm">불러오는 중...</div>
          ) : promisedItems.length === 0 ? (
            <div className="text-center py-16 text-neutral-400 text-sm">약속 대기 중인 고객이 없습니다</div>
          ) : (
            <div className="bg-white rounded-xl border border-neutral-100 overflow-hidden">
              <div className="hidden md:grid grid-cols-[1fr_90px_90px_90px_120px] gap-3 px-4 py-2.5 bg-neutral-50 text-[11px] font-semibold text-neutral-500">
                <div>고객 / 식별번호</div>
                <div>종류</div>
                <div>약속일</div>
                <div>발송</div>
                <div className="text-right">액션</div>
              </div>
              {promisedItems.map((it) => {
                const detailHref = it.source === 'consultation' ? `/consultations/${it.id}` : it.source === 'repair' ? `/repairs/${it.id}` : `/sales/${it.id}`;
                const promisedDate = new Date(it.promisedAt).toLocaleDateString('ko-KR');
                const sentDate = it.requestSentAt ? new Date(it.requestSentAt).toLocaleDateString('ko-KR') : null;
                const related = relatedByItem[`${it.source}-${it.id}`] || [];
                return (
                  <div key={`${it.source}-${it.id}`} className="flex flex-col gap-1.5 md:grid md:grid-cols-[1fr_90px_90px_90px_120px] md:gap-3 md:items-center px-4 py-3 border-t border-neutral-100">
                    <div>
                      <Link href={detailHref} className="hover:underline block">
                        <div className="text-sm font-semibold text-indigo-black">{it.customerName}</div>
                        <div className="text-[11px] text-neutral-500">{it.displayId}</div>
                      </Link>
                      {related.length > 0 && (
                        <div className="flex flex-wrap gap-1 mt-1.5">
                          {related.slice(0, 2).map((r) => {
                            const rHref = r.source === 'consultation' ? `/consultations/${r.id}` : r.source === 'repair' ? `/repairs/${r.id}` : `/sales/${r.id}`;
                            let label = '';
                            let cls = '';
                            if (r.submittedAt) { label = `${r.typeLabel} ✅ 작성완료`; cls = 'bg-green-50 text-green-700 hover:bg-green-100'; }
                            else if (r.requestSentAt) { label = `${r.typeLabel} 📤 발송됨`; cls = 'bg-blue-50 text-blue-700 hover:bg-blue-100'; }
                            else if (r.promisedAt) { label = `${r.typeLabel} ☑ 약속만`; cls = 'bg-amber-50 text-amber-700 hover:bg-amber-100'; }
                            return (
                              <Link
                                key={`${r.source}-${r.id}`}
                                href={rHref}
                                title={`같은 고객 다른 활동 — ${r.displayId}`}
                                className={`text-[10px] px-1.5 py-0.5 rounded font-semibold transition ${cls}`}
                              >
                                🔗 {label}
                              </Link>
                            );
                          })}
                          {related.length > 2 && (
                            <span className="text-[10px] text-neutral-400 px-1 self-center">+{related.length - 2}</span>
                          )}
                        </div>
                      )}
                    </div>
                    <span className="text-[11px] px-2 py-0.5 rounded-full bg-neutral-100 text-neutral-600 inline-block w-fit">{it.typeLabel}</span>
                    <span className="text-xs text-neutral-600">{promisedDate}</span>
                    <span className={`text-[11px] px-2 py-0.5 rounded-full inline-block w-fit ${sentDate ? 'bg-blue-50 text-blue-700' : 'bg-neutral-100 text-neutral-400'}`}>
                      {sentDate ?? '미발송'}
                    </span>
                    <div className="flex justify-end gap-1.5">
                      <button
                        onClick={async () => {
                          if (!window.confirm(`${it.customerName}님 후기 요청 ${sentDate ? '재발송' : '발송'}하시겠습니까?`)) return;
                          const reviewType = it.source === 'consultation' ? 'consult' : it.source === 'repair' ? 'repair' : 'purchase';
                          const res = await fetch('/api/reviews/request', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ source: it.source, id: it.id, review_type: reviewType }),
                          });
                          if (res.ok) {
                            alert('발송 완료');
                            fetchReviews();
                          } else {
                            const e = await res.json().catch(() => ({}));
                            alert(`발송 실패: ${e.error || '알 수 없음'}`);
                          }
                        }}
                        className="px-2 py-1 rounded text-[11px] font-semibold bg-indigo-black text-cream hover:bg-indigo-black/85"
                      >
                        {sentDate ? '재발송' : '발송'}
                      </button>
                      <button
                        onClick={async () => {
                          if (!window.confirm(`${it.customerName}님 약속을 취소하시겠습니까?`)) return;
                          const res = await fetch('/api/reviews/promise', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ source: it.source, id: it.id, on: false }),
                          });
                          if (res.ok) fetchReviews();
                          else alert('취소 실패');
                        }}
                        className="px-2 py-1 rounded text-[11px] text-neutral-500 hover:text-red-500 hover:bg-red-50"
                      >
                        취소
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          )
        )}

        {/* 리뷰 카드 리스트 (promised 외 탭) */}
        {activeTab !== 'promised' && (loading ? (
          <div className="text-center py-16 text-neutral-400 text-sm">불러오는 중...</div>
        ) : reviews.length === 0 ? (
          <div className="text-center py-16 text-neutral-400 text-sm">리뷰가 없습니다</div>
        ) : (
          <div className="grid gap-3 lg:grid-cols-3 md:grid-cols-2">
            {reviews.map(review => (
              <div
                key={review.id}
                className="bg-white rounded-xl border border-neutral-100 p-4 space-y-3"
              >
                {/* 헤더: 이름 + 유형 칩 + 상태 칩 */}
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-sm">{review.name}</span>
                    <span className="text-[10px] font-medium px-2 py-0.5 rounded-full bg-neutral-100 text-neutral-500">
                      {getChipLabel(review)}
                    </span>
                    {getSubChip(review) && (
                      <span className="text-[10px] font-medium px-2 py-0.5 rounded-full bg-neutral-50 text-neutral-400">
                        {getSubChip(review)}
                      </span>
                    )}
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
                      onClick={() => openEdit(review)}
                      className="flex items-center gap-1 px-3 py-1 rounded-lg text-xs font-medium bg-blue-50 text-blue-600 hover:bg-blue-100 transition"
                      title="수정"
                    >
                      <Pencil size={12} /> 수정
                    </button>
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
                    <button
                      onClick={() => setDeleteTarget(review.id)}
                      className="flex items-center gap-1 px-2 py-1 rounded-lg text-xs font-medium bg-red-50 text-red-500 hover:bg-red-100 transition"
                      title="삭제"
                    >
                      <Trash2 size={12} />
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        ))}
      </div>

      {/* ━━━ 수정 모달 ━━━ */}
      {editReview && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-start justify-center p-4 pt-16 overflow-y-auto">
          <div
            className="bg-white rounded-xl w-full max-w-lg p-5 space-y-4"
            onClick={e => e.stopPropagation()}
          >
            <div className="flex items-center justify-between">
              <h3 className="font-semibold text-sm">리뷰 수정</h3>
              <button onClick={() => setEditReview(null)} className="text-neutral-400 hover:text-neutral-600">
                <X size={18} />
              </button>
            </div>

            {/* 유형 */}
            <div>
              <label className="block text-xs font-medium text-neutral-500 mb-1.5">리뷰 유형</label>
              <div className="flex gap-2">
                {TYPE_OPTIONS.map(opt => (
                  <button
                    key={opt.value}
                    onClick={() => setEditForm(prev => ({ ...prev, type: opt.value as Review['type'] }))}
                    className={`px-3 py-1.5 rounded-full text-xs font-medium transition ${
                      editForm.type === opt.value
                        ? 'bg-neutral-800 text-white'
                        : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                    }`}
                  >
                    {opt.label}
                  </button>
                ))}
              </div>
            </div>

            {/* 이름 + 별점 */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">작성자 이름</label>
                <input
                  type="text"
                  value={editForm.name || ''}
                  onChange={e => setEditForm(prev => ({ ...prev, name: e.target.value }))}
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">별점</label>
                <div className="flex gap-1 pt-1">
                  {[1, 2, 3, 4, 5].map(n => (
                    <button
                      key={n}
                      onClick={() => setEditForm(prev => ({ ...prev, stars: n }))}
                      className="transition hover:scale-110"
                    >
                      <Star
                        size={20}
                        className={n <= (editForm.stars || 5) ? 'fill-current text-neutral-800' : 'text-neutral-300'}
                      />
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* 작성일 + 제품명 */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">작성일</label>
                <input
                  type="date"
                  value={editForm.created_at?.slice(0, 10) || ''}
                  onChange={e => setEditForm(prev => ({ ...prev, created_at: e.target.value }))}
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">제품명 (선택)</label>
                <input
                  type="text"
                  value={editForm.product || ''}
                  onChange={e => setEditForm(prev => ({ ...prev, product: e.target.value }))}
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400"
                />
              </div>
            </div>

            {/* 내용 */}
            <div>
              <label className="block text-xs font-medium text-neutral-500 mb-1.5">리뷰 내용</label>
              <textarea
                value={editForm.content || ''}
                onChange={e => setEditForm(prev => ({ ...prev, content: e.target.value }))}
                rows={5}
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400 resize-y"
              />
            </div>

            {/* 사진 */}
            <div>
              <label className="block text-xs font-medium text-neutral-500 mb-1.5">사진</label>
              <div className="flex flex-wrap gap-2">
                {(editForm.photo_urls || []).map((url, i) => (
                  <div key={i} className="relative w-20 h-20 rounded-lg overflow-hidden border border-neutral-100 group">
                    <img src={url} alt="" className="w-full h-full object-cover" />
                    <button
                      onClick={() => setEditForm(prev => ({
                        ...prev,
                        photo_urls: (prev.photo_urls || []).filter((_, idx) => idx !== i),
                      }))}
                      className="absolute top-1 right-1 w-7 h-7 rounded-full bg-black/60 opacity-100 md:opacity-70 md:hover:opacity-100 transition flex items-center justify-center"
                    >
                      <Trash2 size={13} className="text-white" />
                    </button>
                  </div>
                ))}
                <button
                  onClick={() => editPhotoRef.current?.click()}
                  disabled={photoUploading}
                  className="w-20 h-20 rounded-lg border-2 border-dashed border-neutral-200 flex flex-col items-center justify-center text-neutral-400 hover:border-neutral-400 hover:text-neutral-600 transition"
                >
                  {photoUploading ? (
                    <span className="text-xs">업로드중...</span>
                  ) : (
                    <>
                      <ImageIcon size={18} />
                      <span className="text-[10px] mt-1">추가</span>
                    </>
                  )}
                </button>
              </div>
              <input
                ref={editPhotoRef}
                type="file"
                accept="image/jpeg,image/png,image/webp"
                multiple
                className="hidden"
                onChange={handleEditPhotoUpload}
              />
            </div>

            {/* 베스트 */}
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={editForm.is_best || false}
                onChange={e => setEditForm(prev => ({ ...prev, is_best: e.target.checked }))}
                className="rounded border-neutral-300"
              />
              <span className="text-sm text-neutral-600">베스트 리뷰</span>
            </label>

            {/* 저장 버튼 */}
            <div className="flex gap-2 pt-2">
              <button
                onClick={() => setEditReview(null)}
                className="flex-1 py-2.5 rounded-lg border border-neutral-200 text-sm font-medium text-neutral-600 hover:bg-neutral-50 transition"
              >
                취소
              </button>
              <button
                onClick={saveEdit}
                disabled={saving}
                className="flex-1 py-2.5 rounded-lg bg-neutral-800 text-white text-sm font-medium hover:bg-neutral-700 disabled:opacity-50 transition"
              >
                {saving ? '저장 중...' : '저장'}
              </button>
            </div>
          </div>
        </div>
      )}

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

      {/* 삭제 확인 모달 */}
      {deleteTarget && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4" onClick={() => setDeleteTarget(null)}>
          <div className="bg-white rounded-xl p-5 w-[320px] space-y-3" onClick={e => e.stopPropagation()}>
            <h3 className="text-sm font-bold text-neutral-800">리뷰 삭제</h3>
            <p className="text-xs text-neutral-500">이 리뷰를 삭제하시겠습니까? 삭제된 리뷰는 복구할 수 없습니다.</p>
            <div className="flex gap-2">
              <button onClick={() => setDeleteTarget(null)} className="flex-1 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600">취소</button>
              <button onClick={() => handleDelete(deleteTarget)} className="flex-1 py-2 rounded-lg bg-red-500 text-white text-sm font-medium">삭제</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
