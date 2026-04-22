'use client';

/**
 * 아임웹 배너/팝업 원격 관리 UI
 * 설정 > 알림·연동 탭에 삽입
 *
 * MVP: 메인 모달 배너 1종 (id='main_modal')
 * - 이미지 업로드
 * - 제목/설명/링크
 * - 노출 기간 (선택)
 * - 쿠키 시간
 * - 노출 On/Off 토글
 * - 아임웹 주입 스크립트 태그 복사
 */

import { useEffect, useRef, useState } from 'react';
import { Button } from '@/components/ui/button';
import { Megaphone, Image as ImageIcon, Eye, Copy, CheckCircle2, Upload, Zap, AlertCircle } from 'lucide-react';
import toast from 'react-hot-toast';
import { useBanners, useUpdateBanner, useUploadBannerImage, type ImwebBanner } from '@/hooks/use-banner';

const BANNER_ID = 'main_modal';

export default function BannerSettings() {
  const { data: banners, isLoading } = useBanners();
  const update = useUpdateBanner();
  const upload = useUploadBannerImage();

  const [form, setForm] = useState({
    enabled: false,
    title: '',
    description: '',
    image_url: '' as string | null,
    image_path: '' as string | null,
    link_url: '',
    starts_at: '',
    ends_at: '',
    dismiss_cookie_hours: 24,
  });
  const [showPreview, setShowPreview] = useState(false);
  const [copied, setCopied] = useState(false);
  const fileRef = useRef<HTMLInputElement | null>(null);

  const main = (banners || []).find((b) => b.id === BANNER_ID);

  // 초기 로드 시 폼에 반영
  useEffect(() => {
    if (!main) return;
    setForm({
      enabled: main.enabled,
      title: main.title || '',
      description: main.description || '',
      image_url: main.image_url,
      image_path: main.image_path,
      link_url: main.link_url || '',
      starts_at: main.starts_at ? main.starts_at.slice(0, 16) : '',
      ends_at: main.ends_at ? main.ends_at.slice(0, 16) : '',
      dismiss_cookie_hours: main.dismiss_cookie_hours || 24,
    });
  }, [main]);

  const handleSave = () => {
    // 종료일이 과거인데 노출 On이면 경고 — 사장님 실수 예방
    if (form.enabled && form.ends_at) {
      const endDate = new Date(form.ends_at);
      if (!isNaN(endDate.getTime()) && endDate.getTime() < Date.now()) {
        const ok = window.confirm(
          `⚠️ 종료일시가 과거로 설정되어 있습니다.\n\n` +
          `종료일시: ${endDate.toLocaleString('ko-KR')}\n` +
          `현재:     ${new Date().toLocaleString('ko-KR')}\n\n` +
          `이 상태로 저장하면 배너가 고객에게 노출되지 않습니다.\n\n` +
          `계속 저장하시겠습니까?\n` +
          `(취소하고 종료일시를 비우거나 미래로 수정해주세요)`
        );
        if (!ok) return;
      }
    }

    // 시작일이 종료일보다 늦어도 경고
    if (form.starts_at && form.ends_at) {
      const s = new Date(form.starts_at);
      const e = new Date(form.ends_at);
      if (!isNaN(s.getTime()) && !isNaN(e.getTime()) && s >= e) {
        toast.error('시작일시가 종료일시보다 늦거나 같습니다');
        return;
      }
    }

    update.mutate({
      id: BANNER_ID,
      enabled: form.enabled,
      title: form.title || null,
      description: form.description || null,
      link_url: form.link_url || null,
      starts_at: form.starts_at ? new Date(form.starts_at).toISOString() : null,
      ends_at: form.ends_at ? new Date(form.ends_at).toISOString() : null,
      dismiss_cookie_hours: form.dismiss_cookie_hours,
    } as Partial<ImwebBanner> & { id: string });
  };

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const result = await upload.mutateAsync({
        file,
        bannerId: BANNER_ID,
        oldPath: form.image_path || undefined,
      });
      // 즉시 DB에 image_url/image_path 저장 (폼 상태만 바꾸면 저장 전에 소실됨)
      await update.mutateAsync({
        id: BANNER_ID,
        image_url: result.url,
        image_path: result.path,
      });
      setForm((f) => ({ ...f, image_url: result.url, image_path: result.path }));
      toast.success('이미지 업로드 완료');
    } finally {
      if (fileRef.current) fileRef.current.value = '';
    }
  };

  const scriptTag = `<script src="${typeof window !== 'undefined' ? window.location.origin : ''}/api/imweb/banner-widget.js" defer></script>`;

  const handleCopyScript = async () => {
    try {
      await navigator.clipboard.writeText(scriptTag);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
      toast.success('스크립트 태그가 복사되었습니다');
    } catch {
      toast.error('복사 실패 — 수동으로 선택해서 복사해주세요');
    }
  };

  if (isLoading) {
    return (
      <div className="rounded-lg border border-neutral-200 bg-warm-ivory p-4">
        <div className="text-sm text-neutral-500">배너 설정 로드 중...</div>
      </div>
    );
  }

  const saving = update.isPending;
  const uploading = upload.isPending;

  return (
    <div className="rounded-lg border border-neutral-200 bg-warm-ivory p-4 space-y-4">
      {/* 헤더 */}
      <div className="flex items-start gap-3">
        <Megaphone size={20} className="text-neutral-700 mt-0.5 shrink-0" />
        <div className="flex-1">
          <h3 className="text-sm font-bold text-indigo-black">📢 아임웹 배너/팝업 관리</h3>
          <p className="text-xs text-neutral-500 mt-0.5">
            TMS에서 이미지와 텍스트를 설정하면 아임웹 사이트에 모달 배너로 즉시 표시됩니다.
            {main?.updated_at && (
              <>
                <br />
                <span className="text-neutral-400">
                  최종 수정: {new Date(main.updated_at).toLocaleString('ko-KR')}
                </span>
              </>
            )}
          </p>
        </div>
        {/* 노출 중/숨김 뱃지 */}
        <div className={`px-2 py-1 rounded text-[11px] font-bold shrink-0 ${form.enabled ? 'bg-green-100 text-green-700' : 'bg-neutral-200 text-neutral-600'}`}>
          {form.enabled ? '● 노출 중' : '○ 숨김'}
        </div>
      </div>

      {/* 본문 */}
      <div className="rounded-lg bg-white border border-neutral-200 p-4 space-y-4">
        <h4 className="text-sm font-bold text-neutral-700">메인 모달 배너</h4>

        {/* 이미지 업로드 */}
        <div>
          <label className="block text-xs font-semibold text-neutral-700 mb-1.5">이미지</label>
          {form.image_url ? (
            <div className="relative rounded-lg overflow-hidden border border-neutral-200 bg-neutral-50">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src={form.image_url} alt="배너 미리보기" className="w-full max-h-60 object-contain" />
              <button
                onClick={() => fileRef.current?.click()}
                disabled={uploading}
                className="absolute bottom-2 right-2 px-2.5 py-1.5 rounded bg-black/70 text-white text-xs hover:bg-black/90 disabled:opacity-50"
              >
                <Upload size={12} className="inline mr-1" />
                이미지 교체
              </button>
            </div>
          ) : (
            <button
              onClick={() => fileRef.current?.click()}
              disabled={uploading}
              className="w-full h-32 rounded-lg border-2 border-dashed border-neutral-300 bg-neutral-50 hover:bg-neutral-100 flex flex-col items-center justify-center gap-2 text-neutral-500 disabled:opacity-50"
            >
              <ImageIcon size={24} />
              <span className="text-xs">{uploading ? '업로드 중...' : '이미지 선택 (jpg/png/webp, 최대 2MB)'}</span>
            </button>
          )}
          <input ref={fileRef} type="file" accept="image/jpeg,image/png,image/webp" className="hidden" onChange={handleFile} />
        </div>

        {/* 제목 */}
        <Field label="제목">
          <input
            value={form.title}
            onChange={(e) => setForm({ ...form, title: e.target.value })}
            placeholder="예: 사이트 공사중"
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
            maxLength={60}
          />
        </Field>

        {/* 설명 */}
        <Field label="설명 (선택)">
          <textarea
            value={form.description}
            onChange={(e) => setForm({ ...form, description: e.target.value })}
            placeholder="예: 더 나은 서비스를 위해 잠시 공사중입니다."
            rows={3}
            className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm resize-none"
            maxLength={300}
          />
        </Field>

        {/* 링크 */}
        <Field label="클릭 링크 (선택)" desc="배너 이미지 클릭 시 새 탭으로 이동">
          <input
            value={form.link_url}
            onChange={(e) => setForm({ ...form, link_url: e.target.value })}
            placeholder="https://..."
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs"
          />
        </Field>

        {/* 노출 기간 */}
        <div>
          <div className="flex items-center justify-between mb-1.5">
            <label className="block text-xs font-semibold text-neutral-700">노출 기간 (선택)</label>
            {(form.starts_at || form.ends_at) && (
              <button
                type="button"
                onClick={() => setForm({ ...form, starts_at: '', ends_at: '' })}
                className="text-[11px] text-neutral-500 hover:text-neutral-700 underline"
              >
                기간 지우기 (무기한 노출)
              </button>
            )}
          </div>
          <p className="text-[11px] text-neutral-400 mb-1.5">
            💡 비워두면 <b>무기한 노출</b>. 특정 기간만 노출하려면 날짜 입력.
          </p>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-[11px] text-neutral-500 mb-1">시작</label>
              <input
                type="datetime-local"
                value={form.starts_at}
                onChange={(e) => setForm({ ...form, starts_at: e.target.value })}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
              />
            </div>
            <div>
              <label className="block text-[11px] text-neutral-500 mb-1">종료</label>
              <input
                type="datetime-local"
                value={form.ends_at}
                onChange={(e) => setForm({ ...form, ends_at: e.target.value })}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
              />
            </div>
          </div>
        </div>

        {/* 쿠키 시간 */}
        <Field label="닫기 후 재노출 대기 (시간)" desc="'오늘 하루 보지 않기' 클릭 시 이 시간 동안 재노출 안 함">
          <input
            type="number"
            min={0}
            max={720}
            value={form.dismiss_cookie_hours}
            onChange={(e) => setForm({ ...form, dismiss_cookie_hours: Math.max(0, parseInt(e.target.value || '0', 10)) })}
            className="w-32 h-9 px-3 rounded-lg border border-neutral-200 text-sm"
          />
        </Field>

        {/* 노출 토글 */}
        <div className="flex items-center justify-between pt-2 border-t border-neutral-100">
          <div>
            <div className="text-sm font-semibold text-neutral-800">노출 활성화</div>
            <div className="text-xs text-neutral-500">끄면 아임웹 사이트에서 즉시 숨겨집니다 (1분 이내)</div>
          </div>
          <Toggle checked={form.enabled} onChange={(v) => setForm({ ...form, enabled: v })} />
        </div>
      </div>

      {/* 액션 버튼 */}
      <div className="flex flex-wrap gap-2">
        <Button size="sm" onClick={handleSave} disabled={saving || uploading} loading={saving}>
          저장
        </Button>
        <Button size="sm" variant="secondary" onClick={() => setShowPreview(true)} disabled={!form.image_url && !form.title}>
          <Eye size={13} />
          미리보기
        </Button>
      </div>

      {/* 아임웹 자동 설치 (Script API) + 수동 설치 안내 */}
      <ScriptInstallSection scriptTag={scriptTag} onCopyScript={handleCopyScript} copied={copied} />

      {/* 미리보기 모달 */}
      {showPreview && (
        <PreviewModal
          onClose={() => setShowPreview(false)}
          imageUrl={form.image_url || ''}
          title={form.title}
          description={form.description}
        />
      )}
    </div>
  );
}

function ScriptInstallSection({
  scriptTag,
  onCopyScript,
  copied,
}: {
  scriptTag: string;
  onCopyScript: () => void;
  copied: boolean;
}) {
  const [unitCode, setUnitCode] = useState('');
  const [position, setPosition] = useState<'header' | 'body' | 'footer'>('footer');
  const [installing, setInstalling] = useState(false);
  const [installedAt, setInstalledAt] = useState<string>('');
  const [showManual, setShowManual] = useState(false);

  // 현재 설치 상태 로드
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch('/api/imweb/script-install');
        const json = await res.json();
        if (json.ok && json.data) {
          if (json.data.unit_code) setUnitCode(json.data.unit_code);
          if (json.data.position) setPosition(json.data.position);
          if (json.data.installed_at) setInstalledAt(json.data.installed_at);
        }
      } catch { /* 무시 */ }
    })();
  }, []);

  const handleAutoInstall = async () => {
    if (!unitCode.trim()) {
      toast.error('unitCode를 먼저 입력해주세요');
      return;
    }
    setInstalling(true);
    try {
      const res = await fetch('/api/imweb/script-install', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ unitCode: unitCode.trim(), position }),
      });
      const json = await res.json();
      if (!res.ok || !json.ok) {
        if (json.needsReauth) {
          toast.error('아임웹 OAuth 재연결 필요 — script:write 권한이 없습니다');
        } else {
          toast.error('설치 실패: ' + (json.error || 'unknown'));
        }
        return;
      }
      const action = json.data.action;
      setInstalledAt(new Date().toISOString());
      toast.success(`아임웹에 자동 설치 완료 (${action === 'created' ? '신규 등록' : '업데이트'})`);
    } finally {
      setInstalling(false);
    }
  };

  return (
    <div className="rounded-lg bg-neutral-50 border border-neutral-200 p-3 space-y-3">
      <div className="flex items-center gap-2">
        <p className="text-xs font-semibold text-neutral-700">📌 아임웹 설치 (최초 1회만)</p>
      </div>

      {/* 자동 설치 블록 */}
      <div className="rounded border border-neutral-200 bg-white p-3 space-y-2.5">
        <div className="flex items-center gap-1.5">
          <Zap size={13} className="text-amber-600" />
          <span className="text-xs font-bold text-neutral-800">자동 설치 (Script API)</span>
          {installedAt && (
            <span className="ml-auto text-[10px] text-green-700 bg-green-50 px-1.5 py-0.5 rounded">
              ✓ 설치됨 {new Date(installedAt).toLocaleDateString('ko-KR')}
            </span>
          )}
        </div>

        <div className="grid grid-cols-[1fr_auto] gap-2">
          <input
            value={unitCode}
            onChange={(e) => setUnitCode(e.target.value)}
            placeholder="unitCode (예: u20250402e3b6987310679)"
            className="h-8 px-2.5 rounded border border-neutral-200 text-xs font-mono"
          />
          <select
            value={position}
            onChange={(e) => setPosition(e.target.value as 'header' | 'body' | 'footer')}
            className="h-8 px-2 rounded border border-neutral-200 text-xs"
          >
            <option value="footer">footer (권장)</option>
            <option value="body">body</option>
            <option value="header">header</option>
          </select>
        </div>

        <div className="flex items-center gap-2">
          <Button size="sm" onClick={handleAutoInstall} disabled={installing || !unitCode.trim()} loading={installing}>
            <Zap size={13} />
            {installedAt ? '재설치' : '자동 설치'}
          </Button>
          <button onClick={() => setShowManual(!showManual)} className="text-[11px] text-neutral-500 hover:text-neutral-700 underline">
            {showManual ? '수동 설치 숨기기' : '수동 설치 방법 보기'}
          </button>
        </div>

        <p className="text-[10px] text-neutral-400 leading-relaxed">
          💡 unitCode는 아임웹 관리자 → 사이트 관리 → URL 또는 콘솔에서 확인 가능.
          position은 footer(권장)를 선택하면 페이지 로딩 지연 없음.
        </p>
      </div>

      {/* 수동 설치 블록 (접히는 영역) */}
      {showManual && (
        <div className="rounded border border-neutral-200 bg-white p-3 space-y-2">
          <div className="flex items-center gap-1.5">
            <AlertCircle size={13} className="text-neutral-500" />
            <span className="text-xs font-bold text-neutral-800">수동 설치 (백업용)</span>
          </div>
          <p className="text-[11px] text-neutral-500 leading-relaxed">
            자동 설치 안 되면 아임웹 관리자 → 디자인 → 코드위젯에 아래 태그를 붙여넣어 주세요.
          </p>
          <div className="flex items-center gap-1.5 bg-neutral-50 border border-neutral-200 rounded p-2">
            <code className="flex-1 text-[11px] font-mono text-neutral-700 break-all">{scriptTag}</code>
            <button onClick={onCopyScript} className="shrink-0 p-1.5 rounded hover:bg-neutral-100" title="복사">
              {copied ? <CheckCircle2 size={14} className="text-green-600" /> : <Copy size={14} className="text-neutral-500" />}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-neutral-700 mb-1">{label}</label>
      {desc && <p className="text-[11px] text-neutral-400 mb-1.5">{desc}</p>}
      {children}
    </div>
  );
}

function Toggle({ checked, onChange }: { checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <button
      type="button"
      onClick={() => onChange(!checked)}
      className={`relative w-12 h-7 rounded-full transition ${checked ? 'bg-green-600' : 'bg-neutral-300'}`}
    >
      <span className={`absolute top-0.5 left-0.5 w-6 h-6 rounded-full bg-white shadow transition-transform ${checked ? 'translate-x-5' : ''}`} />
    </button>
  );
}

function PreviewModal({
  onClose,
  imageUrl,
  title,
  description,
}: {
  onClose: () => void;
  imageUrl: string;
  title: string;
  description: string;
}) {
  return (
    <div
      className="fixed inset-0 z-[9999] bg-black/70 flex items-center justify-center p-6"
      onClick={onClose}
    >
      <div
        className="bg-[#FAF9F7] max-w-[420px] w-full rounded-lg overflow-hidden shadow-2xl relative"
        onClick={(e) => e.stopPropagation()}
      >
        <button
          onClick={onClose}
          className="absolute top-2.5 right-2.5 w-8 h-8 rounded-full bg-[#FAF9F7]/90 hover:bg-white flex items-center justify-center z-10"
        >
          ×
        </button>
        {imageUrl && (
          // eslint-disable-next-line @next/next/no-img-element
          <img src={imageUrl} alt="preview" className="w-full block" />
        )}
        {(title || description) && (
          <div className="p-6">
            {title && <h3 className="text-[17px] font-bold text-[#1A1A1A] mb-2 -tracking-[0.3px]">{title}</h3>}
            {description && <p className="text-sm text-[#555] leading-relaxed whitespace-pre-line">{description}</p>}
          </div>
        )}
        <div className="flex border-t border-[#EEE]">
          <button className="flex-1 py-3.5 text-xs text-[#555] hover:bg-[#F5F5F5]">오늘 하루 보지 않기</button>
          <button className="flex-1 py-3.5 text-xs font-bold text-[#1A1A1A] border-l border-[#EEE] hover:bg-[#F5F5F5]">닫기</button>
        </div>
      </div>
    </div>
  );
}
