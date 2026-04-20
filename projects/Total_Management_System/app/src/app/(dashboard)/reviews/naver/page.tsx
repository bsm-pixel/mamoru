'use client';

import { useState, useRef, useCallback } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Upload, FileText, Plus, Trash2, Star, ImageIcon, X, Check, AlertCircle } from 'lucide-react';

type ReviewType = 'consult' | 'repair' | 'purchase';

interface NaverReview {
  type: ReviewType;
  subtype: string;
  name: string;
  stars: number;
  content: string;
  photo_urls: string[];
  product: string;
  is_best: boolean;
  created_at: string;
  received_at: string; // 상담일/접수일/구매일
}

const EMPTY_REVIEW: NaverReview = {
  type: 'consult',
  subtype: 'field_request',
  name: '',
  stars: 5,
  content: '',
  photo_urls: [],
  product: '',
  is_best: false,
  created_at: '',
  received_at: '',
};

const TYPE_OPTIONS: { value: ReviewType; label: string }[] = [
  { value: 'consult', label: '상담' },
  { value: 'repair', label: '복원수리' },
  { value: 'purchase', label: '제품구매' },
];

const CONSULT_SUBTYPES: { value: string; label: string }[] = [
  { value: 'field_request', label: '출장' },
  { value: 'store_visit', label: '직접방문' },
];

export default function NaverReviewPage() {
  const [mode, setMode] = useState<'single' | 'csv'>('single');
  const [review, setReview] = useState<NaverReview>({ ...EMPTY_REVIEW });
  const [csvReviews, setCsvReviews] = useState<NaverReview[]>([]);
  const [uploading, setUploading] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [result, setResult] = useState<{ success: number; failed: number } | null>(null);
  const [photoUploading, setPhotoUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const csvInputRef = useRef<HTMLInputElement>(null);
  const photoInputRef = useRef<HTMLInputElement>(null);

  /* ── 사진 업로드 ── */
  const uploadPhotos = useCallback(async (files: FileList): Promise<string[]> => {
    setPhotoUploading(true);
    try {
      const { resizeImage } = await import('@/lib/utils/resize-image');
      const resized = await Promise.all(Array.from(files).map(f => resizeImage(f, 1200, 0.8)));
      const formData = new FormData();
      resized.forEach(f => formData.append('files', f));

      const res = await fetch('/api/reviews/upload-bulk', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);

      return (data.results || [])
        .filter((r: { url?: string }) => r.url)
        .map((r: { url: string }) => r.url);
    } catch (err) {
      console.error('사진 업로드 실패:', err);
      alert('사진 업로드 실패: ' + String(err));
      return [];
    } finally {
      setPhotoUploading(false);
    }
  }, []);

  /* ── 단건 등록 ── */
  const submitSingle = async () => {
    if (!review.name || !review.content) {
      alert('이름과 내용은 필수입니다');
      return;
    }
    setSubmitting(true);
    setResult(null);
    try {
      const res = await fetch('/api/reviews/naver', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reviews: [review] }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setResult({ success: data.success, failed: data.failed });
      if (data.success > 0) setReview({ ...EMPTY_REVIEW });
    } catch (err) {
      alert('등록 실패: ' + String(err));
    } finally {
      setSubmitting(false);
    }
  };

  /* ── CSV 파싱 ── */
  const handleCsvUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const text = evt.target?.result as string;
        const lines = text.split('\n').filter(l => l.trim());
        if (lines.length < 2) {
          alert('CSV 헤더 + 최소 1행 필요');
          return;
        }

        // 헤더: type,name,stars,content,created_at,received_at,product,subtype,is_best
        const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
        const reviews: NaverReview[] = [];

        for (let i = 1; i < lines.length; i++) {
          const values = parseCsvLine(lines[i]);
          const row: Record<string, string> = {};
          headers.forEach((h, idx) => { row[h] = (values[idx] || '').trim(); });

          reviews.push({
            type: (['consult', 'repair', 'purchase'].includes(row.type) ? row.type : 'consult') as ReviewType,
            name: row.name || '',
            stars: Math.min(5, Math.max(1, parseInt(row.stars) || 5)),
            content: row.content || '',
            created_at: row.created_at || '',
            received_at: row.received_at || '',
            product: row.product || '',
            subtype: row.subtype || '',
            is_best: row.is_best === 'true' || row.is_best === '1',
            photo_urls: [], // 사진은 별도 매핑
          });
        }

        setCsvReviews(reviews);
      } catch (err) {
        alert('CSV 파싱 실패: ' + String(err));
      } finally {
        setUploading(false);
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  /* ── CSV 일괄 제출 ── */
  const submitCsv = async () => {
    if (csvReviews.length === 0) return;
    setSubmitting(true);
    setResult(null);
    try {
      // 50건씩 배치 처리
      let totalSuccess = 0;
      let totalFailed = 0;

      for (let i = 0; i < csvReviews.length; i += 50) {
        const batch = csvReviews.slice(i, i + 50);
        const res = await fetch('/api/reviews/naver', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ reviews: batch }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);
        totalSuccess += data.success;
        totalFailed += data.failed;
      }

      setResult({ success: totalSuccess, failed: totalFailed });
      if (totalSuccess > 0) setCsvReviews([]);
    } catch (err) {
      alert('일괄 등록 실패: ' + String(err));
    } finally {
      setSubmitting(false);
    }
  };

  /* ── 단건 사진 추가 ── */
  const handlePhotoSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const urls = await uploadPhotos(files);
    setReview(prev => ({ ...prev, photo_urls: [...prev.photo_urls, ...urls] }));
    e.target.value = '';
  };

  const removePhoto = (idx: number) => {
    setReview(prev => ({
      ...prev,
      photo_urls: prev.photo_urls.filter((_, i) => i !== idx),
    }));
  };

  return (
    <>
      <Topbar title="네이버 리뷰 등록" />

      <div className="px-4 md:px-6 py-4 space-y-6 max-w-3xl">
        {/* 모드 전환 */}
        <div className="flex gap-2">
          <button
            onClick={() => setMode('single')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-medium transition ${
              mode === 'single' ? 'bg-neutral-800 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
            }`}
          >
            <Plus size={14} /> 단건 등록
          </button>
          <button
            onClick={() => setMode('csv')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-medium transition ${
              mode === 'csv' ? 'bg-neutral-800 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
            }`}
          >
            <FileText size={14} /> CSV 일괄 등록
          </button>
        </div>

        {/* 결과 알림 */}
        {result && (
          <div className={`flex items-center gap-2 px-4 py-3 rounded-lg text-sm font-medium ${
            result.failed === 0
              ? 'bg-green-50 text-green-700'
              : 'bg-yellow-50 text-yellow-700'
          }`}>
            <Check size={16} />
            {result.success}건 등록 완료{result.failed > 0 && ` / ${result.failed}건 실패`}
          </div>
        )}

        {/* ━━━ 단건 등록 폼 ━━━ */}
        {mode === 'single' && (
          <div className="bg-white rounded-xl border border-neutral-100 p-5 space-y-4">
            {/* 유형 */}
            <div>
              <label className="block text-xs font-medium text-neutral-500 mb-1.5">리뷰 유형</label>
              <div className="flex gap-2">
                {TYPE_OPTIONS.map(opt => (
                  <button
                    key={opt.value}
                    onClick={() => setReview(prev => ({ ...prev, type: opt.value, subtype: opt.value === 'consult' ? 'field_request' : '' }))}
                    className={`px-3 py-1.5 rounded-full text-xs font-medium transition ${
                      review.type === opt.value
                        ? 'bg-neutral-800 text-white'
                        : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                    }`}
                  >
                    {opt.label}
                  </button>
                ))}
              </div>
            </div>

            {/* 상담 방식 (consult 선택 시) */}
            {review.type === 'consult' && (
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">상담 방식</label>
                <div className="flex gap-2">
                  {CONSULT_SUBTYPES.map(opt => (
                    <button
                      key={opt.value}
                      onClick={() => setReview(prev => ({ ...prev, subtype: opt.value }))}
                      className={`px-3 py-1.5 rounded-full text-xs font-medium transition ${
                        review.subtype === opt.value
                          ? 'bg-neutral-800 text-white'
                          : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* 이름 + 별점 */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">작성자 이름</label>
                <input
                  type="text"
                  value={review.name}
                  onChange={e => setReview(prev => ({ ...prev, name: e.target.value }))}
                  placeholder="홍길동"
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">별점</label>
                <div className="flex gap-1 pt-1">
                  {[1, 2, 3, 4, 5].map(n => (
                    <button
                      key={n}
                      onClick={() => setReview(prev => ({ ...prev, stars: n }))}
                      className="transition hover:scale-110"
                    >
                      <Star
                        size={20}
                        className={n <= review.stars ? 'fill-current text-neutral-800' : 'text-neutral-300'}
                      />
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* 작성일 + 상담일/접수일/구매일 */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">리뷰 작성일</label>
                <input
                  type="date"
                  value={review.created_at}
                  onChange={e => setReview(prev => ({ ...prev, created_at: e.target.value }))}
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-neutral-500 mb-1.5">
                  {review.type === 'repair' ? '접수일' : review.type === 'purchase' ? '구매일' : '상담일'}
                </label>
                <input
                  type="date"
                  value={review.received_at}
                  onChange={e => setReview(prev => ({ ...prev, received_at: e.target.value }))}
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400"
                />
              </div>
            </div>

            {/* 제품명 */}
            <div>
              <label className="block text-xs font-medium text-neutral-500 mb-1.5">제품명 (선택)</label>
              <input
                type="text"
                value={review.product}
                onChange={e => setReview(prev => ({ ...prev, product: e.target.value }))}
                placeholder={review.type === 'purchase' ? '구매한 제품명 (예: MAMORU M7)' : '관련 제품명'}
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400"
              />
            </div>

            {/* 내용 */}
            <div>
              <label className="block text-xs font-medium text-neutral-500 mb-1.5">리뷰 내용</label>
              <textarea
                value={review.content}
                onChange={e => setReview(prev => ({ ...prev, content: e.target.value }))}
                rows={5}
                placeholder="네이버 플레이스 리뷰 내용을 붙여넣으세요"
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:border-neutral-400 resize-y"
              />
            </div>

            {/* 사진 */}
            <div>
              <label className="block text-xs font-medium text-neutral-500 mb-1.5">사진</label>
              <div className="flex flex-wrap gap-2">
                {review.photo_urls.map((url, i) => (
                  <div key={i} className="relative w-20 h-20 rounded-lg overflow-hidden border border-neutral-100 group">
                    <img src={url} alt="" className="w-full h-full object-cover" />
                    <button
                      onClick={() => removePhoto(i)}
                      className="absolute inset-0 bg-black/50 opacity-0 group-hover:opacity-100 transition flex items-center justify-center"
                    >
                      <Trash2 size={16} className="text-white" />
                    </button>
                  </div>
                ))}
                <button
                  onClick={() => photoInputRef.current?.click()}
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
                ref={photoInputRef}
                type="file"
                accept="image/jpeg,image/png,image/webp"
                multiple
                className="hidden"
                onChange={handlePhotoSelect}
              />
            </div>

            {/* 베스트 리뷰 */}
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={review.is_best}
                onChange={e => setReview(prev => ({ ...prev, is_best: e.target.checked }))}
                className="rounded border-neutral-300"
              />
              <span className="text-sm text-neutral-600">베스트 리뷰로 지정</span>
            </label>

            {/* 등록 버튼 */}
            <button
              onClick={submitSingle}
              disabled={submitting}
              className="w-full py-2.5 rounded-lg bg-neutral-800 text-white text-sm font-medium hover:bg-neutral-700 disabled:opacity-50 transition"
            >
              {submitting ? '등록 중...' : '리뷰 등록'}
            </button>
          </div>
        )}

        {/* ━━━ CSV 일괄 등록 ━━━ */}
        {mode === 'csv' && (
          <div className="space-y-4">
            {/* CSV 가이드 */}
            <div className="bg-neutral-50 rounded-xl border border-neutral-100 p-4 space-y-3">
              <div className="flex items-start gap-2">
                <AlertCircle size={16} className="text-neutral-500 mt-0.5 shrink-0" />
                <div className="text-sm text-neutral-600 space-y-1">
                  <p className="font-medium">CSV 형식 안내</p>
                  <p className="text-xs text-neutral-500">
                    첫 행은 헤더, 이후 행이 리뷰 데이터입니다.
                  </p>
                  <code className="block text-xs bg-white p-2 rounded border border-neutral-200 overflow-x-auto">
                    type,name,stars,content,created_at,product,subtype,is_best
                  </code>
                  <ul className="text-xs text-neutral-500 space-y-0.5 list-disc pl-4">
                    <li>type: consult / repair / purchase</li>
                    <li>stars: 1~5</li>
                    <li>created_at: YYYY-MM-DD (네이버 작성일)</li>
                    <li>is_best: true 또는 1 (선택)</li>
                    <li>content에 쉼표가 있으면 &quot;큰따옴표&quot;로 감싸기</li>
                  </ul>
                </div>
              </div>
            </div>

            {/* CSV 파일 선택 */}
            <button
              onClick={() => csvInputRef.current?.click()}
              disabled={uploading}
              className="w-full py-10 rounded-xl border-2 border-dashed border-neutral-200 flex flex-col items-center gap-2 text-neutral-400 hover:border-neutral-400 hover:text-neutral-600 transition"
            >
              <Upload size={24} />
              <span className="text-sm font-medium">
                {uploading ? '파일 읽는 중...' : 'CSV 파일 선택'}
              </span>
              <span className="text-xs">.csv 파일을 선택하세요</span>
            </button>
            <input
              ref={csvInputRef}
              type="file"
              accept=".csv"
              className="hidden"
              onChange={handleCsvUpload}
            />

            {/* CSV 미리보기 */}
            {csvReviews.length > 0 && (
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-sm font-medium text-neutral-700">
                    {csvReviews.length}건 파싱 완료
                  </span>
                  <button
                    onClick={() => setCsvReviews([])}
                    className="text-xs text-neutral-400 hover:text-neutral-600"
                  >
                    <X size={14} /> 초기화
                  </button>
                </div>

                <div className="max-h-80 overflow-auto rounded-xl border border-neutral-100">
                  <table className="w-full text-xs">
                    <thead className="bg-neutral-50 sticky top-0">
                      <tr>
                        <th className="px-3 py-2 text-left font-medium text-neutral-500">#</th>
                        <th className="px-3 py-2 text-left font-medium text-neutral-500">유형</th>
                        <th className="px-3 py-2 text-left font-medium text-neutral-500">이름</th>
                        <th className="px-3 py-2 text-left font-medium text-neutral-500">별점</th>
                        <th className="px-3 py-2 text-left font-medium text-neutral-500">내용</th>
                        <th className="px-3 py-2 text-left font-medium text-neutral-500">날짜</th>
                      </tr>
                    </thead>
                    <tbody>
                      {csvReviews.slice(0, 20).map((r, i) => (
                        <tr key={i} className="border-t border-neutral-50">
                          <td className="px-3 py-2 text-neutral-400">{i + 1}</td>
                          <td className="px-3 py-2">{r.type}</td>
                          <td className="px-3 py-2 font-medium">{r.name}</td>
                          <td className="px-3 py-2">{'★'.repeat(r.stars)}</td>
                          <td className="px-3 py-2 max-w-[200px] truncate">{r.content}</td>
                          <td className="px-3 py-2 text-neutral-400">{r.created_at}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {csvReviews.length > 20 && (
                    <div className="text-center py-2 text-xs text-neutral-400">
                      외 {csvReviews.length - 20}건...
                    </div>
                  )}
                </div>

                <button
                  onClick={submitCsv}
                  disabled={submitting}
                  className="w-full py-2.5 rounded-lg bg-neutral-800 text-white text-sm font-medium hover:bg-neutral-700 disabled:opacity-50 transition"
                >
                  {submitting ? '등록 중...' : `${csvReviews.length}건 일괄 등록`}
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </>
  );
}

/** CSV 라인 파싱 (큰따옴표 내 쉼표 처리) */
function parseCsvLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}
