// Vercel ignoreCommand 최종 테스트 — TMS 빌드 진행 확인
import { redirect } from 'next/navigation';

export default function Home() {
  redirect('/dashboard');
}
