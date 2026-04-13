// Vercel Ignored Build Step 테스트 v2 — TMS 빌드 진행 확인
import { redirect } from 'next/navigation';

export default function Home() {
  redirect('/dashboard');
}
