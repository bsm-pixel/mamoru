// Vercel 테스트 v6 — --quiet 추가 후 TMS 빌드 확인
import { redirect } from 'next/navigation';

export default function Home() {
  redirect('/dashboard');
}
