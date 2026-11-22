/**
 * contracts.image_url 에 든 값 → 'contracts' 저장소 안 위치 (2026-11-22)
 *   새 값은 위치 그대로('contracts/<id>.png'). 옛 공개 주소(…/object/public/contracts/<path>)도 위치로 바꿔 읽는다.
 *   계약서 이미지는 api/contracts/[id]/image GET 이 로그인 확인 뒤 1분짜리 주소로 내준다.
 */
export function contractImagePath(imageUrl: string | null | undefined): string | null {
  const v = (imageUrl || '').trim();
  if (!v) return null;
  const m = v.match(/\/object\/(?:public|sign)\/contracts\/([^?]+)/);
  if (m) return decodeURIComponent(m[1]);
  if (/^https?:\/\//.test(v)) return null;
  return v;
}
