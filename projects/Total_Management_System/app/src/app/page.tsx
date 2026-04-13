// Vercel Ignored Build Step 테스트 — 이 커밋은 빌드 진행되어야 함
import { redirect } from 'next/navigation';

export default function Home() {
  redirect('/dashboard');
}
