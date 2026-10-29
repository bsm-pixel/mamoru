/**
 * 메일 영수증 원본 보관 (2026-10-29)
 *
 * 사장님: 해외 결제(Anthropic·Vercel·Make 등)는 홈택스에 안 잡혀 「영수증만」 증빙 → 원본을 따로 보관해야 한다.
 * 구글 드라이브는 콘솔 설정·재연결이 필요해, TMS 파일 저장소(Supabase Storage 비공개 버킷)에 둔다 — 사장님 손 0.
 *
 * 대상: 메일에서 확인한 경비(source_ref 'mail:…') 중 증빙이 영수증만·없음·미정이고 아직 보관 안 된 것.
 *       (사업자카드·세금계산서·현금영수증은 홈택스 자동이라 제외)
 * 저장: PDF 첨부가 있으면 그것(Stripe 영수증+인보이스 등 전부), 없으면 메일 본문 HTML 1개.
 *       경로 YYYY/MM/날짜_거래처_금액_메일id[_n].pdf|html
 * 기록: expenses.receipt_url = 'storage:' + 경로들('|' 구분). 보기 = GET /api/expenses/receipt?id=&i= (로그인 → 1분짜리 서명 주소)
 * 호출: 대기함 [확인] 직후(after) + 매일 크론(못 한 것·옛 것 재시도, 회당 최대 30건).
 */
import { getMailFiles } from '@/lib/google/gmail-client';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export const RECEIPT_BUCKET = 'expense-receipts';

const slug = (s: string) => s.normalize('NFKD').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 30) || 'vendor';

export function receiptPaths(receiptUrl: string | null | undefined): string[] {
  if (!receiptUrl?.startsWith('storage:')) return [];
  return receiptUrl.slice('storage:'.length).split('|').filter(Boolean);
}

async function ensureBucket(db: Db): Promise<void> {
  const { data } = await db.storage.getBucket(RECEIPT_BUCKET);
  if (data) return;
  const { error } = await db.storage.createBucket(RECEIPT_BUCKET, { public: false, fileSizeLimit: 20 * 1024 * 1024 });
  if (error && !/already exists/i.test(error.message)) throw error;
}

interface ExpenseRow { id: string; expense_date: string; amount: number; vendor: string | null; source_ref: string | null }

/** 경비 1건 보관. 반환: 저장한 경로들 (메일을 못 읽으면 throw 대신 error) */
export async function archiveExpenseReceipt(db: Db, e: ExpenseRow): Promise<{ paths?: string[]; error?: string }> {
  const gmailId = e.source_ref?.startsWith('mail:') ? e.source_ref.slice(5) : null;
  if (!gmailId) return { error: '메일 경비 아님' };
  const mail = await getMailFiles(gmailId);
  if (!mail.ok) return { error: mail.needsReauth ? '구글 재연결 필요' : mail.notConnected ? '구글 미연결' : (mail.error || '메일 읽기 실패') };

  const files: Array<{ data: Buffer; ext: 'pdf' | 'html'; type: string }> = mail.pdfs.length
    ? mail.pdfs.map((p) => ({ data: p.data, ext: 'pdf' as const, type: 'application/pdf' }))
    : mail.html
      ? [{
          ext: 'html' as const, type: 'text/html; charset=utf-8',
          // 보관용 머리말 — 보낸 곳·제목·받은 시각 (원본 본문은 그대로 아래에)
          data: Buffer.from(`<!doctype html><meta charset="utf-8"><div style="font:12px sans-serif;color:#555;border-bottom:1px solid #ccc;padding:8px;margin-bottom:12px">보낸 곳: ${(mail.from || '').replace(/</g, '&lt;')}<br>제목: ${(mail.subject || '').replace(/</g, '&lt;')}<br>받은 시각: ${mail.receivedAt}<br>TMS 경비 ${e.expense_date} · ${e.amount.toLocaleString()}원</div>${mail.html}`, 'utf-8'),
        }]
      : [];
  if (!files.length) return { error: '메일에 영수증 내용이 없음' };

  await ensureBucket(db);
  const [y, m] = e.expense_date.split('-');
  const base = `${y}/${m}/${e.expense_date}_${slug(e.vendor || '')}_${e.amount}_${gmailId}`;
  const paths: string[] = [];
  for (let i = 0; i < files.length; i++) {
    const path = `${base}${files.length > 1 ? `_${i + 1}` : ''}.${files[i].ext}`;
    const { error } = await db.storage.from(RECEIPT_BUCKET).upload(path, files[i].data, { contentType: files[i].type, upsert: true });
    if (error) throw error;
    paths.push(path);
  }
  const { error: upErr } = await db.from('expenses').update({ receipt_url: `storage:${paths.join('|')}` }).eq('id', e.id).is('receipt_url', null);
  if (upErr) throw upErr;
  return { paths };
}

/** 보관 안 된 대상 경비를 최대 limit 건 보관 (크론·확인 직후) */
export async function archivePendingReceipts(db: Db, limit = 30, onlyId?: string): Promise<{ archived: number; failed: Array<{ id: string; error: string }> }> {
  let q = db.from('expenses').select('id, expense_date, amount, vendor, source_ref, proof_type')
    .like('source_ref', 'mail:%').is('receipt_url', null)
    .or('proof_type.is.null,proof_type.in.(receipt,none)')
    .order('expense_date', { ascending: false }).limit(limit);
  if (onlyId) q = q.eq('id', onlyId);
  const { data, error } = await q;
  if (error) throw error;
  let archived = 0;
  const failed: Array<{ id: string; error: string }> = [];
  for (const e of data || []) {
    const r = await archiveExpenseReceipt(db, e);
    if (r.paths) archived++;
    else failed.push({ id: e.id, error: r.error || '실패' });
    // 구글 연결 문제면 나머지도 똑같이 실패 — 바로 멈춤
    if (r.error && /구글/.test(r.error)) break;
  }
  return { archived, failed };
}
