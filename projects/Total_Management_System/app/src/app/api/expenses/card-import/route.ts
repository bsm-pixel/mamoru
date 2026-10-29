import { NextRequest, NextResponse } from 'next/server';
import { requireUser } from '@/lib/auth/require-user';
import { createServiceClient } from '@/lib/supabase/server';
import { parseCardSheet, importCardRows } from '@/lib/expenses/card-import';

/**
 * POST /api/expenses/card-import (multipart: file) — 카드 이용내역 엑셀 → 경비 「확인 대기」 (2026-10-29)
 * 규칙은 lib/expenses/card-import.ts. 같은 파일 두 번 올려도 한 번만 들어간다.
 */
export async function POST(req: NextRequest) {
  const deny = await requireUser(); if (deny) return deny;
  try {
    const form = await req.formData();
    const file = form.get('file');
    if (!(file instanceof File)) return NextResponse.json({ error: '파일을 골라 주세요' }, { status: 400 });
    if (file.size > 10 * 1024 * 1024) return NextResponse.json({ error: '파일이 너무 큽니다 (10MB 이하)' }, { status: 400 });

    const parsed = parseCardSheet(await file.arrayBuffer());
    if (parsed.error) return NextResponse.json({ error: parsed.error }, { status: 400 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await importCardRows(createServiceClient() as any, parsed.rows, file.name);
    return NextResponse.json({ ok: true, rows: parsed.rows.length, cancelled: parsed.cancelled, ...result });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
