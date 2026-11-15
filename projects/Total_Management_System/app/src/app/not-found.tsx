import Link from 'next/link';

// 2026-11-15 없는 주소로 들어왔을 때 — 전엔 영어 기본 화면만 나오고 돌아갈 길이 없었다
export default function NotFound() {
  return (
    <div className="min-h-screen flex items-center justify-center bg-stone-50 p-6">
      <div className="text-center">
        <p className="text-sm font-bold text-stone-900">이 주소에는 화면이 없습니다</p>
        <p className="mt-1.5 text-xs text-neutral-500">주소가 바뀌었거나 잘못 들어온 것일 수 있습니다.</p>
        <Link href="/dashboard" className="mt-5 inline-flex items-center justify-center h-10 px-5 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800">
          대시보드로
        </Link>
      </div>
    </div>
  );
}
