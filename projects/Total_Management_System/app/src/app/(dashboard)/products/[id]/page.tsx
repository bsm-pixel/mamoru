import { redirect } from 'next/navigation';

// /products/[id] 는 도달 경로가 없는 유령 라우트였음(제품 상세는 목록의 ProductDetailPanel 모달이 SSOT).
// 직접 URL 접근 시 404 대신 제품 목록으로 보낸다. (서브라우트 /products/[id]/serials 는 그대로 유지)
export default function ProductIdRedirect() {
  redirect('/products');
}
