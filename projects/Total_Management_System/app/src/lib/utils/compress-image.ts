/**
 * 업로드 전 이미지 압축 — 수리내역서 v3 (가위 사진·확대 사진)
 *
 * - square: 가운데 기준 1:1 로 잘라낸 뒤 maxSize×maxSize (가위 사진 — 고객 화면도 1:1)
 * - 그 외: 긴 변 maxSize 로 비율 유지 (확대 사진 — 현미경 원본 비율 그대로)
 * - WebP 우선. 브라우저가 WebP 인코딩을 못 하면(iOS 사파리 등 toBlob 이 PNG 로 떨어짐) JPEG 로 저장
 * - 작은 이미지도 항상 다시 인코딩한다 (기존 resizeImage 는 작으면 원본 PNG/HEIC 를 그대로 올렸음)
 */

export interface CompressOptions {
  maxSize?: number;
  quality?: number;
  square?: boolean;
}

function loadImage(file: Blob): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('이미지를 읽을 수 없습니다')); };
    img.src = url;
  });
}

function toBlob(canvas: HTMLCanvasElement, type: string, quality: number): Promise<Blob | null> {
  return new Promise((resolve) => canvas.toBlob(resolve, type, quality));
}

export async function compressImage(file: File, opts: CompressOptions = {}): Promise<File> {
  const { maxSize = 1080, quality = 0.82, square = false } = opts;
  const img = await loadImage(file);
  const w = img.naturalWidth, h = img.naturalHeight;

  // 원본에서 잘라낼 영역
  let sx = 0, sy = 0, sw = w, sh = h;
  if (square) { const side = Math.min(w, h); sx = (w - side) / 2; sy = (h - side) / 2; sw = sh = side; }

  const scale = Math.min(1, maxSize / Math.max(sw, sh));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(sw * scale);
  canvas.height = Math.round(sh * scale);
  const ctx = canvas.getContext('2d');
  if (!ctx) return file;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(img, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height);

  let blob = await toBlob(canvas, 'image/webp', quality);
  if (!blob || blob.type !== 'image/webp') blob = await toBlob(canvas, 'image/jpeg', quality);
  if (!blob) return file;
  const ext = blob.type === 'image/webp' ? 'webp' : 'jpg';
  const base = file.name.replace(/\.[^.]+$/, '') || 'photo';
  return new File([blob], `${base}.${ext}`, { type: blob.type, lastModified: Date.now() });
}
