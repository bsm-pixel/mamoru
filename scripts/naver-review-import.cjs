#!/usr/bin/env node
/**
 * 네이버 스마트플레이스 리뷰 → TMS 일괄 등록
 *
 * 입력: projects/marketing/naver_review_extract.js 로 내려받아 압축 푼 폴더
 *       (reviews.json + 고객별 폴더(review.md, photo_XX.jpg) — 사진은 download_images.ps1 실행 후)
 *
 * 2단계로 쓴다 (종류 판단은 사람이/Claude가 본문을 읽고 확정한다 — 네이버 파일엔 서비스명이 비어 있다):
 *   1) node scripts/naver-review-import.cjs plan  "<폴더>"
 *        → 이미 등록된 건(중복) 제외하고 <폴더>/_tms_plan.json 생성. 각 건에 규칙 기반 '추정 종류'가 들어 있다.
 *   2) 본문을 읽고 종류를 확정해 <폴더>/_tms_overrides.json 에 적는다 → plan 을 다시 실행하면 반영된다
 *        { "<index>": { "type": "repair", "subtype": "direct_visit" },
 *          "<index>": { "skip": true, "skip_reason": "사장님 답글만 있음" } }
 *   3) node scripts/naver-review-import.cjs apply "<폴더>"
 *        → 사진 업로드 + reviews 등록. 다시 실행해도 중복 등록되지 않는다.
 *
 * 종류(type) / 세부(subtype)
 *   consult  : store_visit(매장 방문 상담) | field_request(출장 상담) | talk_consult(전화·톡 상담)
 *   repair   : direct_visit(매장 방문 수리) | parcel_pickup(택배·수거 수리) | restoration(구분 불명)
 *   purchase : (subtype 없음)
 *
 * 저장 규칙은 TMS 화면(api/reviews/naver)과 같다: 별점 5(네이버엔 별점 없음), status approved, source 'naver',
 *   created_at = 네이버 작성일, meta.received_at = 방문일.
 */
const fs = require('fs');
const path = require('path');

const APP = path.join(__dirname, '..', 'projects', 'Total_Management_System', 'app');
for (const l of fs.readFileSync(path.join(APP, '.env.local'), 'utf8').split(/\r?\n/)) {
  const m = l.match(/^([A-Z0-9_]+)=(.*)$/);
  if (m) process.env[m[1]] = m[2].replace(/^"|"$/g, '');
}
const { createClient } = require(path.join(APP, 'node_modules', '@supabase', 'supabase-js'));
const db = createClient(process.env.NEXT_PUBLIC_SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);

const [mode, dirArg] = process.argv.slice(2);
const DIR = dirArg || path.join(process.env.USERPROFILE || '', 'Desktop', '네이버리뷰 정리 폴더');
const PLAN = path.join(DIR, '_tms_plan.json');
const norm = (s) => String(s || '').replace(/[\s\p{P}\p{S}]/gu, '').slice(0, 30);

/** 네이버 화면 잔재(업체명·작성시각·수정/삭제 버튼 글자) 제거 */
function cleanContent(raw) {
  return String(raw || '').split('\n').map((l) => l.trim()).filter((t) =>
    t && t !== '수정' && t !== '삭제' && t !== '완료' && t !== '마모루 미용가위'
    && !/^\d{4}\.\s*\d+\.\s*\d+\s*\(\S+\)\s*(오전|오후)/.test(t)).join('\n').replace(/\s*접기$/, '').trim();
}

/** 규칙 기반 추정 — 확정이 아니다. plan 에서 사람이/Claude가 본문을 읽고 고친다 */
function guess(c) {
  const has = (re) => re.test(c);
  const repair = has(/복원|수리|연마|날\s?갈|갈아|고쳐|수선|a\/?s|에이에스|틴닝.*(맞|조절)|볼트/i);
  const buy = has(/구매|구입|샀|사게|장만|주문|결제|구매했|데려/);
  const consult = has(/상담|컨설팅|추천|설명|알려주|골라|비교/);
  const field = has(/출장|와주|방문해\s?주|찾아와\s?주|매장으로\s?와|샵으로\s?와|오셔서/);
  const parcel = has(/택배|보내|수거|배송/);
  let type = 'consult', subtype = field ? 'field_request' : 'store_visit', sure = false;
  if (repair && !buy) { type = 'repair'; subtype = parcel ? 'parcel_pickup' : 'direct_visit'; sure = !consult; }
  else if (buy && !repair) { type = 'purchase'; subtype = null; sure = true; }
  else if (repair && buy) { type = 'repair'; subtype = parcel ? 'parcel_pickup' : 'direct_visit'; }
  else if (consult) sure = true;
  return { type, subtype, sure, why: [repair && '수리', buy && '구매', consult && '상담', field && '출장', parcel && '택배'].filter(Boolean).join('+') || '단서 없음' };
}

async function existingNaver() {
  const { data, error } = await db.from('reviews').select('review_id, created_at, name, content, meta').eq('source', 'naver');
  if (error) throw error;
  return (data || []).map((r) => ({ date: (r.meta?.naver_date || r.created_at || '').slice(0, 10), key: norm(r.content), id: r.review_id }));
}

async function plan() {
  const all = JSON.parse(fs.readFileSync(path.join(DIR, 'reviews.json'), 'utf8'));
  const exist = await existingNaver();
  const items = all.map((r) => {
    const content = cleanContent(r.content);
    const folder = path.join(DIR, r.folder);
    const photos = fs.existsSync(folder) ? fs.readdirSync(folder).filter((f) => /^photo_.*\.(jpe?g|png|webp)$/i.test(f)).sort() : [];
    const videos = fs.existsSync(folder) ? fs.readdirSync(folder).filter((f) => /\.(mp4|mov|webm)$/i.test(f)) : [];
    const dup = exist.find((e) => e.date === r.write_date && e.key && e.key === norm(content));
    const g = guess(content);
    return {
      index: r.index, folder: r.folder, name: (r.reviewer || '').trim() || '네이버 고객',
      write_date: r.write_date, visit_date: r.visit_date || '',
      type: g.type, subtype: g.subtype, guess_sure: g.sure, guess_why: g.why,
      skip: !!dup || !content, skip_reason: dup ? `이미 등록됨(${dup.id})` : (!content ? '본문 없음' : ''),
      photos, expected_photos: (r.photos || []).length, videos: videos.length,
      content,
    };
  });
  // 확정값 덮어쓰기 — 규칙 추정보다 우선. plan 을 다시 돌려도 판정이 유지된다
  const OV = path.join(DIR, '_tms_overrides.json');
  if (fs.existsSync(OV)) {
    const ov = JSON.parse(fs.readFileSync(OV, 'utf8'));
    for (const it of items) {
      const o = ov[String(it.index)];
      if (!o) continue;
      if (o.type) { it.type = o.type; it.subtype = o.subtype ?? null; it.guess_sure = true; it.guess_why = '확정'; }
      if (o.skip && !it.skip) { it.skip = true; it.skip_reason = o.skip_reason || '제외'; }
    }
  }
  fs.writeFileSync(PLAN, JSON.stringify(items, null, 1));
  const t = (f) => items.filter(f).length;
  console.log(`plan 저장: ${PLAN}`);
  console.log(`전체 ${items.length} · 등록 예정 ${t((i) => !i.skip)} · 중복 ${t((i) => /이미/.test(i.skip_reason))} · 본문 없음 ${t((i) => i.skip_reason === '본문 없음')}`);
  console.log(`추정: 상담 ${t((i) => !i.skip && i.type === 'consult')} · 복원수리 ${t((i) => !i.skip && i.type === 'repair')} · 제품구매 ${t((i) => !i.skip && i.type === 'purchase')} · 확신 낮음 ${t((i) => !i.skip && !i.guess_sure)}`);
  console.log(`사진: 폴더에 있는 것 ${items.reduce((s, i) => s + i.photos.length, 0)} / 네이버 기준 ${items.reduce((s, i) => s + i.expected_photos, 0)} · 영상 ${items.reduce((s, i) => s + i.videos, 0)}(영상은 등록하지 않음)`);
}

async function nextReviewId(dateStr, used) {
  const prefix = `RV-${dateStr.replace(/-/g, '')}-`;
  if (!(prefix in used)) {
    const { data } = await db.from('reviews').select('review_id').like('review_id', `${prefix}%`).order('review_id', { ascending: false }).limit(1);
    used[prefix] = data && data.length ? parseInt(data[0].review_id.split('-').pop() || '0', 10) : 0;
  }
  used[prefix] += 1;
  return `${prefix}${String(used[prefix]).padStart(3, '0')}`;
}

async function apply() {
  const items = JSON.parse(fs.readFileSync(PLAN, 'utf8'));
  const exist = await existingNaver();
  const used = {};
  let ok = 0, skipped = 0, failed = 0, photoCount = 0;
  // 오래된 것부터 등록 — 같은 날짜 안에서 리뷰 번호가 작성 순서대로 붙는다
  for (const it of [...items].sort((a, b) => (a.write_date || '').localeCompare(b.write_date || '') || b.index - a.index)) {
    if (it.skip) { skipped++; continue; }
    if (exist.find((e) => e.date === it.write_date && e.key === norm(it.content))) { skipped++; continue; }
    try {
      const urls = [];
      for (const f of it.photos) {
        const ext = path.extname(f).toLowerCase().replace('.', '') || 'jpg';
        // 경로를 리뷰·파일명으로 고정 → 다시 실행해도 같은 자리에 덮어쓴다(중복 파일 안 생김)
        const key = `reviews/naver/${it.write_date}_${String(it.index).padStart(3, '0')}_${f}`;
        const { error: upErr } = await db.storage.from('review-photos').upload(key, fs.readFileSync(path.join(DIR, it.folder, f)), {
          contentType: ext === 'png' ? 'image/png' : ext === 'webp' ? 'image/webp' : 'image/jpeg', upsert: true,
        });
        if (upErr) throw upErr;
        urls.push(db.storage.from('review-photos').getPublicUrl(key).data.publicUrl);
      }
      const created = `${it.write_date}T00:00:00+00:00`;
      const { error } = await db.from('reviews').insert({
        review_id: await nextReviewId(it.write_date, used),
        type: it.type, subtype: it.subtype || null,
        name: it.name, phone: '', stars: 5, content: it.content, photo_urls: urls,
        source_id: `naver-${it.write_date}-${String(it.index).padStart(3, '0')}`,
        status: 'approved', approved_at: new Date().toISOString(), source: 'naver', is_best: false, product: null,
        meta: { naver_date: it.write_date, imported_at: new Date().toISOString(), ...(it.visit_date ? { received_at: it.visit_date } : {}) },
        created_at: created,
      });
      if (error) throw error;
      ok++; photoCount += urls.length;
    } catch (e) {
      failed++;
      console.error(`실패 #${it.index} ${it.write_date} ${it.name}: ${e.message || e}`);
    }
  }
  console.log(`등록 ${ok} · 건너뜀 ${skipped} · 실패 ${failed} · 사진 ${photoCount}`);
}

(mode === 'plan' ? plan() : mode === 'apply' ? apply() : Promise.reject(new Error('사용법: node scripts/naver-review-import.cjs plan|apply "<폴더>"')))
  .catch((e) => { console.error(e.message || e); process.exit(1); });
