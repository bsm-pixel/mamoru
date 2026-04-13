// Vercel Ignored Build Step 테스트 v3 — Custom 모드 빌드 확인
import { redirect } from 'next/navigation';

export default function Home() {
  redirect('/dashboard');
}
