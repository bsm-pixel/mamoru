'use client';

import { Suspense, useState } from 'react';
import { useRouter, useSearchParams } from 'next/navigation';
import { createClient } from '@/lib/supabase/client';
import toast from 'react-hot-toast';

function LoginForm() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const router = useRouter();
  const searchParams = useSearchParams();
  const redirect = searchParams.get('redirect') || '/dashboard';

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);

    const supabase = createClient();
    const { error } = await supabase.auth.signInWithPassword({ email, password });

    if (error) {
      toast.error('로그인 실패: ' + error.message);
      setLoading(false);
      return;
    }

    toast.success('로그인 성공');
    router.push(redirect);
    router.refresh();
  };

  return (
    <form onSubmit={handleLogin} className="bg-card-white rounded-2xl p-6 shadow-sm border border-neutral-200">
      <div className="space-y-4">
        <div>
          <label htmlFor="email" className="block text-sm font-medium text-neutral-700 mb-1">
            이메일
          </label>
          <input
            id="email"
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
            autoComplete="email"
            className="w-full h-11 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 focus:border-terracotta transition"
            placeholder="admin@mamoru.kr"
          />
        </div>
        <div>
          <label htmlFor="password" className="block text-sm font-medium text-neutral-700 mb-1">
            비밀번호
          </label>
          <input
            id="password"
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
            autoComplete="current-password"
            className="w-full h-11 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 focus:border-terracotta transition"
            placeholder="••••••••"
          />
        </div>
      </div>

      <button
        type="submit"
        disabled={loading}
        className="mt-6 w-full h-11 rounded-lg bg-terracotta text-cream font-semibold text-[15px] hover:bg-terracotta-deep active:scale-[0.98] transition disabled:opacity-50 disabled:cursor-not-allowed"
      >
        {loading ? '로그인 중...' : '로그인'}
      </button>
    </form>
  );
}

export default function LoginPage() {
  return (
    <div className="min-h-screen flex items-center justify-center bg-cream px-4">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <h1 className="text-2xl font-extrabold tracking-tight text-indigo-black">
            MAMORU
          </h1>
          <p className="mt-1 text-sm text-neutral-500">통합관리시스템</p>
        </div>

        <Suspense fallback={<div className="h-64 animate-pulse bg-card-white rounded-2xl" />}>
          <LoginForm />
        </Suspense>

        <p className="mt-4 text-center text-xs text-neutral-400">
          MAMORU TMS v1.0
        </p>
      </div>
    </div>
  );
}
