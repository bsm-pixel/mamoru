/* ──────────────────────────────────────────────────────────────
   catalog.js — 카탈로그 자산 로더
   manifest.json 을 읽고 cards / copy_pool JSON 을 전부 fetch 하여
   조회하기 쉬운 인덱스 형태로 정리한다. (JSON 이 SSOT, 런타임 로드)
   ────────────────────────────────────────────────────────────── */
const Catalog = {
  cards: [],          // 카드 정의 배열 (card_type 별)
  copy: [],           // 카피 풀 정의 배열 (copy_type 별, 여러 종류 가능)
  byCardType: {},     // card_type -> 카드 정의
  loaded: false,

  async load() {
    const manifest = await fetchJSON('./manifest.json');
    const base = manifest.catalog_base || '../catalog';

    const cardFiles = manifest.cards || [];
    const copyFiles = manifest.copy_pool || [];

    this.cards = await Promise.all(
      cardFiles.map(p => fetchJSON(`${base}/${p}`).then(j => ({ ...j, _src: p })))
    );
    this.copy = await Promise.all(
      copyFiles.map(p => fetchJSON(`${base}/${p}`).then(j => ({ ...j, _src: p })))
    );

    this.byCardType = {};
    for (const c of this.cards) this.byCardType[c.card_type] = c;

    this.loaded = true;
    return this;
  },

  /* 특정 가위 종류에 적용되는 카드만 (manifest 순서 유지) */
  cardsForType(type) {
    return this.cards.filter(c => (c.applies_to || []).includes(type));
  },

  /* 특정 종류 + copy_type 에 맞는 카피 풀 1개 (가장 구체적인 것 우선) */
  copyPool(copyType, type) {
    const cands = this.copy.filter(
      p => p.copy_type === copyType && (p.applies_to || []).includes(type)
    );
    if (!cands.length) return null;
    // applies_to 가 좁은(전용) 풀을 우선 — 예: blunt 전용 > 전 종류 공용
    cands.sort((a, b) => (a.applies_to || []).length - (b.applies_to || []).length);
    return cands[0];
  },

  /* 카피 풀의 옵션 id -> text 조회 */
  copyText(copyType, type, id) {
    const pool = this.copyPool(copyType, type);
    if (!pool) return id;
    const opt = (pool.options || []).find(o => o.id === id);
    return opt ? opt.text : id;
  },

  /* 카드 옵션 id -> 옵션 객체 */
  cardOption(cardType, id) {
    const card = this.byCardType[cardType];
    if (!card) return null;
    return (card.options || []).find(o => o.id === id) || null;
  }
};

async function fetchJSON(url) {
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`로드 실패 ${url} (${res.status})`);
  return res.json();
}
