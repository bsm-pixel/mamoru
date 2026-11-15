/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * 가짜 자료 저장소 — 수파베이스 조회 문법 중 TMS가 쓰는 부분만 메모리에서 흉내 낸다.
 * 운영 자료에 닿지 않게 하려는 것. 지원하지 않는 문법을 만나면 조용히 넘어가지 않고 오류를 낸다(거짓 통과 방지).
 */
type Row = Record<string, any>;
type Rel = { table: string; fk: string; many: boolean; localKey?: string };

let idSeq = 0;
export const newId = () => `id-${String(++idSeq).padStart(4, '0')}`;

export class FakeDb {
  tables: Record<string, Row[]> = {};
  user: { id: string } | null = { id: 'user-1' };
  rpcCalls: Array<{ name: string; args: any }> = [];
  rpcHandlers: Record<string, (args: any, db: FakeDb) => any> = {};
  /** 부모표 → 끼워 읽는 이름 → 연결 규칙 */
  relations: Record<string, Record<string, Rel>> = {
    offline_sales: { offline_sale_items: { table: 'offline_sale_items', fk: 'sale_id', many: true } },
    purchase_orders: { purchase_order_items: { table: 'purchase_order_items', fk: 'po_id', many: true } },
    deliveries: { delivery_items: { table: 'delivery_items', fk: 'delivery_id', many: true } },
    product_serials: { products: { table: 'products', fk: 'id', many: false, localKey: 'product_id' } },
  };
  defaults: Record<string, Row> = {};
  /** 데이터베이스가 스스로 채우는 칸 흉내(예: 전화번호 숫자만) */
  computed: Record<string, (r: Row) => void> = {
    customers: (r) => { r.phone_normalized = String(r.phone || '').replace(/\D/g, '') || null; },
  };

  auth = { getUser: async () => ({ data: { user: this.user }, error: null }) };
  storage = {
    from: () => ({
      upload: async () => ({ data: {}, error: null }),
      remove: async () => ({ data: [], error: null }),
      getPublicUrl: () => ({ data: { publicUrl: 'https://fake.local/file' } }),
      createSignedUrl: async () => ({ data: { signedUrl: 'https://fake.local/signed' }, error: null }),
    }),
  };

  seed(table: string, rows: Row[]) {
    this.tables[table] = [...(this.tables[table] || []), ...rows.map((r) => ({ ...r }))];
    return this;
  }
  rows(table: string): Row[] { return this.tables[table] || []; }
  one(table: string, id: string): Row | undefined { return this.rows(table).find((r) => r.id === id); }

  from(table: string) { return new Query(this, table); }
  rpc(name: string, args?: any) {
    this.rpcCalls.push({ name, args });
    const h = this.rpcHandlers[name];
    const out = h ? { data: h(args, this), error: null } : { data: null, error: { message: `가짜 저장소: 함수 ${name} 미정의`, code: 'FAKE' } };
    return Promise.resolve(out);
  }
}

const cmp = (a: any, b: any) => (a === b ? 0 : a === null || a === undefined ? -1 : b === null || b === undefined ? 1 : a < b ? -1 : 1);
const likeRe = (pat: string, ci: boolean) =>
  new RegExp('^' + String(pat).replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/%/g, '.*').replace(/_/g, '.') + '$', ci ? 'i' : '');

class Query implements PromiseLike<any> {
  private filters: Array<(r: Row) => boolean> = [];
  private op: 'select' | 'insert' | 'update' | 'delete' | 'upsert' = 'select';
  private payload: any = null;
  private cols = '*';
  private returning = false;
  private mode: 'many' | 'single' | 'maybe' = 'many';
  private orders: Array<{ col: string; asc: boolean }> = [];
  private lim: number | null = null;
  private off = 0;
  private wantCount = false;
  private head = false;
  private upsertKey: string[] = ['id'];

  constructor(private db: FakeDb, private table: string) {}

  select(cols = '*', opts?: { count?: string; head?: boolean }) {
    if (this.op === 'select') { this.cols = cols; } else { this.returning = true; this.cols = cols; }
    if (opts?.count) this.wantCount = true;
    if (opts?.head) this.head = true;
    return this;
  }
  insert(v: any) { this.op = 'insert'; this.payload = v; return this; }
  update(v: any) { this.op = 'update'; this.payload = v; return this; }
  delete() { this.op = 'delete'; return this; }
  upsert(v: any, opts?: { onConflict?: string }) {
    this.op = 'upsert'; this.payload = v;
    if (opts?.onConflict) this.upsertKey = opts.onConflict.split(',').map((s) => s.trim());
    return this;
  }

  eq(c: string, v: any) { this.filters.push((r) => r[c] === v); return this; }
  neq(c: string, v: any) { this.filters.push((r) => r[c] !== v); return this; }
  gt(c: string, v: any) { this.filters.push((r) => r[c] != null && r[c] > v); return this; }
  gte(c: string, v: any) { this.filters.push((r) => r[c] != null && r[c] >= v); return this; }
  lt(c: string, v: any) { this.filters.push((r) => r[c] != null && r[c] < v); return this; }
  lte(c: string, v: any) { this.filters.push((r) => r[c] != null && r[c] <= v); return this; }
  in(c: string, vs: any[]) { this.filters.push((r) => vs.includes(r[c])); return this; }
  is(c: string, v: any) { this.filters.push((r) => (v === null ? r[c] === null || r[c] === undefined : r[c] === v)); return this; }
  like(c: string, p: string) { const re = likeRe(p, false); this.filters.push((r) => re.test(String(r[c] ?? ''))); return this; }
  ilike(c: string, p: string) { const re = likeRe(p, true); this.filters.push((r) => re.test(String(r[c] ?? ''))); return this; }
  match(o: Row) { for (const [k, v] of Object.entries(o)) this.eq(k, v); return this; }
  not(c: string, op: string, v: any) {
    if (op === 'is') { this.filters.push((r) => !(v === null ? r[c] === null || r[c] === undefined : r[c] === v)); return this; }
    if (op === 'eq') { this.filters.push((r) => r[c] !== v); return this; }
    if (op === 'in') {
      const list = String(v).replace(/^\(|\)$/g, '').split(',').map((s) => s.trim().replace(/^"|"$/g, ''));
      this.filters.push((r) => !list.includes(String(r[c]))); return this;
    }
    throw new Error(`가짜 저장소: not(${op}) 미지원 — ${this.table}`);
  }
  or(expr: string) {
    const parts = expr.split(',').map((s) => s.trim()).filter(Boolean);
    const tests = parts.map((p) => {
      const m = /^([a-z_0-9]+)\.(eq|neq|is|ilike|like|gt|gte|lt|lte)\.(.*)$/i.exec(p);
      if (!m) throw new Error(`가짜 저장소: or("${p}") 미지원 — ${this.table}`);
      const [, c, op, raw] = m;
      const v: any = raw === 'null' ? null : raw === 'true' ? true : raw === 'false' ? false : raw;
      if (op === 'eq') return (r: Row) => String(r[c]) === String(v);
      if (op === 'neq') return (r: Row) => String(r[c]) !== String(v);
      if (op === 'is') return (r: Row) => (v === null ? r[c] == null : r[c] === v);
      if (op === 'ilike' || op === 'like') { const re = likeRe(String(v), op === 'ilike'); return (r: Row) => re.test(String(r[c] ?? '')); }
      if (op === 'gt') return (r: Row) => r[c] > v;
      if (op === 'gte') return (r: Row) => r[c] >= v;
      if (op === 'lt') return (r: Row) => r[c] < v;
      return (r: Row) => r[c] <= v;
    });
    this.filters.push((r) => tests.some((t) => t(r)));
    return this;
  }
  order(col: string, opts?: { ascending?: boolean }) { this.orders.push({ col, asc: opts?.ascending !== false }); return this; }
  limit(n: number) { this.lim = n; return this; }
  range(a: number, b: number) { this.off = a; this.lim = b - a + 1; return this; }
  single() { this.mode = 'single'; return this; }
  maybeSingle() { this.mode = 'maybe'; return this; }

  then<T1 = any, T2 = never>(ok?: ((v: any) => T1 | PromiseLike<T1>) | null, bad?: ((e: any) => T2 | PromiseLike<T2>) | null): PromiseLike<T1 | T2> {
    let out: any;
    try { out = this.exec(); } catch (e) { return Promise.reject(e).then(ok as any, bad as any); }
    return Promise.resolve(out).then(ok as any, bad as any);
  }

  private matchRows(): Row[] { return this.db.rows(this.table).filter((r) => this.filters.every((f) => f(r))); }

  private project(row: Row): Row {
    const out: Row = { ...row };
    // 끼워 읽기: name(...) / alias:name(...) / name!fk(...)
    const re = /(?:([a-z_0-9]+):)?([a-z_0-9]+)(?:!([a-z_0-9]+))?\s*\(/gi;
    let m: RegExpExecArray | null;
    while ((m = re.exec(this.cols))) {
      const alias = m[1] || m[2];
      const rel = this.db.relations[this.table]?.[m[2]];
      if (!rel) throw new Error(`가짜 저장소: ${this.table} → ${m[2]} 끼워 읽기 규칙 없음 (fake-db.ts relations 에 추가)`);
      const pool = this.db.rows(rel.table);
      out[alias] = rel.many
        ? pool.filter((c) => c[rel.fk] === row.id).map((c) => ({ ...c }))
        : (pool.find((c) => c[rel.fk] === row[rel.localKey || 'id']) ? { ...pool.find((c) => c[rel.fk] === row[rel.localKey || 'id']) } : null);
    }
    return out;
  }

  private finish(rows: Row[], count?: number) {
    if (this.mode === 'single') {
      if (rows.length !== 1) return { data: null, error: { message: `행 ${rows.length}개 (single)`, code: 'PGRST116' }, count: count ?? null };
      return { data: rows[0], error: null, count: count ?? null };
    }
    if (this.mode === 'maybe') {
      if (rows.length > 1) return { data: null, error: { message: '행 여러 개 (maybeSingle)', code: 'PGRST116' }, count: count ?? null };
      return { data: rows[0] ?? null, error: null, count: count ?? null };
    }
    return { data: this.head ? null : rows, error: null, count: count ?? null };
  }

  private exec() {
    const t = (this.db.tables[this.table] ||= []);
    const now = new Date().toISOString();
    if (this.op === 'insert' || this.op === 'upsert') {
      const list: Row[] = Array.isArray(this.payload) ? this.payload : [this.payload];
      const made: Row[] = [];
      for (const p of list) {
        if (this.op === 'upsert') {
          const ex = t.find((r) => this.upsertKey.every((k) => r[k] !== undefined && r[k] === p[k]));
          if (ex) { Object.assign(ex, p, { updated_at: now }); made.push(ex); continue; }
        }
        const row = { id: newId(), created_at: now, updated_at: now, ...(this.db.defaults[this.table] || {}), ...p };
        this.db.computed[this.table]?.(row);
        t.push(row); made.push(row);
      }
      return this.returning ? this.finish(made.map((r) => this.project(r))) : { data: null, error: null };
    }
    if (this.op === 'update') {
      const hit = this.matchRows();
      for (const r of hit) { Object.assign(r, this.payload); this.db.computed[this.table]?.(r); }
      return this.returning ? this.finish(hit.map((r) => this.project(r))) : { data: null, error: null };
    }
    if (this.op === 'delete') {
      const hit = new Set(this.matchRows());
      this.db.tables[this.table] = t.filter((r) => !hit.has(r));
      return this.returning ? this.finish([...hit].map((r) => this.project(r))) : { data: null, error: null };
    }
    let rows = this.matchRows();
    const total = rows.length;
    for (const o of [...this.orders].reverse()) rows = [...rows].sort((a, b) => (o.asc ? 1 : -1) * cmp(a[o.col], b[o.col]));
    if (this.off) rows = rows.slice(this.off);
    if (this.lim !== null) rows = rows.slice(0, this.lim);
    return this.finish(rows.map((r) => this.project(r)), this.wantCount ? total : undefined);
  }
}
