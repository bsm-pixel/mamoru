/**
 * ══════════════════════════════════════════════════════════════════
 *  MAMORU — 네이버 스마트플레이스 리뷰 일괄 추출 스크립트
 * ══════════════════════════════════════════════════════════════════
 *
 *  사용법:
 *  1) 크롬에서 네이버 스마트플레이스 관리자 로그인
 *  2) 리뷰 관리 페이지 이동 (new.smartplace.naver.com/.../reviews)
 *  3) 모든 리뷰가 로드될 때까지 스크롤 끝까지 내리기
 *     "더보기" 버튼 있으면 끝까지 클릭
 *  4) F12 → Console 탭
 *  5) 이 파일 전체 복사 → Console에 붙여넣기 → Enter
 *     ⚠ "붙여넣기 허용" 경고 뜨면: `allow pasting` 입력 후 Enter, 다시 붙여넣기
 *  6) 자동 실행 → naver_reviews_YYYY-MM-DD.zip 다운로드됨
 *
 *  결과물:
 *    ├─ reviews.json        — 전체 리뷰 메타데이터
 *    ├─ 001_날짜_이름/       — 고객별 폴더
 *    │   ├─ photo_01.jpg
 *    │   └─ photo_02.jpg
 *    └─ _failed.json        — 다운로드 실패 목록
 * ══════════════════════════════════════════════════════════════════
 */

(async function() {
  'use strict';

  console.log('%c[MAMORU] 네이버 리뷰 추출 시작', 'font-size:16px;font-weight:bold;color:#03C75A');

  // ─────────────────────────────────────────────────────────────
  // 0. JSZip 동적 로드 (폴더 구조 ZIP 생성용)
  // ─────────────────────────────────────────────────────────────
  if (typeof JSZip === 'undefined') {
    console.log('📦 JSZip 라이브러리 로드 중...');
    await new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
    console.log('✅ JSZip 로드 완료');
  }

  // ─────────────────────────────────────────────────────────────
  // 1. 접힌 리뷰 본문 모두 펼치기 ("더보기" 버튼 자동 클릭)
  // ─────────────────────────────────────────────────────────────
  console.log('🔎 접힌 리뷰 본문 펼치는 중...');
  let expanded = 0;
  document.querySelectorAll('button, a, span').forEach(el => {
    const txt = (el.textContent || '').trim();
    if (txt === '더보기' || txt === '펼치기') {
      try { el.click(); expanded++; } catch {}
    }
  });
  if (expanded > 0) {
    console.log(`✅ ${expanded}개 본문 펼침 — DOM 안정화 대기`);
    await new Promise(r => setTimeout(r, 800));
  }

  // ─────────────────────────────────────────────────────────────
  // 2. 리뷰 카드 찾기 (구조 기반 셀렉터)
  //    네이버 클래스명이 난독화되어 있으므로 내부 구조로 판별
  // ─────────────────────────────────────────────────────────────
  function findReviewCards() {
    // 후보 셀렉터들 (사이트 업데이트 대비 여러 패턴)
    const candidates = [
      'li[class*="review"]',
      'div[class*="review_item"]',
      'div[class*="ReviewItem"]',
      'ul[class*="review"] > li',
    ];
    for (const sel of candidates) {
      const els = document.querySelectorAll(sel);
      if (els.length >= 1) {
        // 리뷰 특징: 본문 또는 작성일 텍스트 포함
        const filtered = [...els].filter(el => {
          const t = el.textContent || '';
          return t.includes('작성일') || t.includes('방문일') || /\d{4}\.\s*\d{1,2}\.\s*\d{1,2}/.test(t);
        });
        if (filtered.length > 0) return filtered;
      }
    }
    return [];
  }

  const cards = findReviewCards();
  if (cards.length === 0) {
    console.error('❌ 리뷰 카드를 찾지 못했습니다. 페이지 구조가 변경되었거나 로그인 상태를 확인하세요.');
    alert('리뷰 카드를 찾지 못했습니다. Console 로그를 확인해주세요.');
    return;
  }
  console.log(`📋 리뷰 ${cards.length}건 발견`);

  // ─────────────────────────────────────────────────────────────
  // 3. 유틸 함수
  // ─────────────────────────────────────────────────────────────
  const DATE_RE = /(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/;
  const TIME_RE = /(오[전후])\s*(\d{1,2}):(\d{2})/;

  function parseDate(text) {
    const m = text.match(DATE_RE);
    if (!m) return '';
    return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  }

  function parseTime(text) {
    const m = text.match(TIME_RE);
    if (!m) return '';
    let hour = parseInt(m[2]);
    if (m[1] === '오후' && hour < 12) hour += 12;
    if (m[1] === '오전' && hour === 12) hour = 0;
    return `${String(hour).padStart(2,'0')}:${m[3]}`;
  }

  function safeFolder(s) {
    return (s || '')
      .replace(/[<>:"/\\|?*]/g, '_')
      .replace(/\s+/g, '_')
      .substring(0, 40);
  }

  async function fetchBlob(url) {
    try {
      const res = await fetch(url, { credentials: 'include' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return await res.blob();
    } catch (e) {
      return null;
    }
  }

  // URL에서 확장자 추정 (네이버 CDN은 확장자 없이 쿼리로 제공되는 경우 많음)
  function guessExt(url, blob) {
    const m = url.match(/\.(jpg|jpeg|png|webp|gif|mp4|mov|webm)(?:\?|$)/i);
    if (m) return m[1].toLowerCase();
    if (blob?.type) {
      if (blob.type.includes('png')) return 'png';
      if (blob.type.includes('webp')) return 'webp';
      if (blob.type.includes('gif')) return 'gif';
      if (blob.type.includes('mp4')) return 'mp4';
      if (blob.type.includes('webm')) return 'webm';
    }
    return 'jpg';
  }

  // ─────────────────────────────────────────────────────────────
  // 4. 각 카드에서 데이터 추출
  // ─────────────────────────────────────────────────────────────
  function extractCard(el, idx) {
    const text = el.textContent || '';

    // 리뷰어 이름 (마스킹된 형태: 김*관, 이*희 등)
    // 페이지 안에서 *를 포함한 짧은 문자열을 찾는다
    let reviewer = '';
    const reviewerMatch = text.match(/([가-힣])\*+([가-힣]?)/);
    if (reviewerMatch) reviewer = reviewerMatch[0];

    // 작성일
    const writeDateEl = [...el.querySelectorAll('*')].find(n => {
      const t = (n.textContent || '').trim();
      return t.startsWith('작성일') && t.length < 30;
    });
    const writeDate = writeDateEl ? parseDate(writeDateEl.textContent) : '';

    // 방문일 (+시간)
    const visitEl = [...el.querySelectorAll('*')].find(n => {
      const t = (n.textContent || '').trim();
      return t.startsWith('방문일') && t.length < 50;
    });
    const visitDate = visitEl ? parseDate(visitEl.textContent) : '';
    const visitTime = visitEl ? parseTime(visitEl.textContent) : '';

    // 예약자/완료 상태
    const statusMatch = text.match(/예약자[^·]*?(완료|취소|미방문|노쇼)\s*(\d+)?/);
    const status = statusMatch ? statusMatch[1] : '';

    // 업체명 · 서비스명 (예: "마모루 미용가위 · 가위 컨설팅상담 (마모루 사무실 방문)")
    let business = '', service = '';
    const businessLine = [...el.querySelectorAll('*')].find(n => {
      const t = (n.textContent || '').trim();
      return /마모루.*·/.test(t) && t.length < 100 && !n.querySelector('*');
    });
    if (businessLine) {
      const parts = businessLine.textContent.split('·').map(s => s.trim());
      business = parts[0] || '';
      service = parts.slice(1).join(' · ') || '';
    }

    // 본문 — 가장 긴 텍스트 블록 (대체로 p/div 중 제일 긴 것)
    let content = '';
    const textNodes = [...el.querySelectorAll('p, div, span')]
      .map(n => ({ el: n, txt: (n.innerText || '').trim() }))
      .filter(x => x.txt.length > 30 && !x.txt.includes('작성일') && !x.txt.includes('방문일'))
      .sort((a,b) => b.txt.length - a.txt.length);
    if (textNodes.length > 0) content = textNodes[0].txt;

    // 태그 칩 (품질이 좋아요, A/S가 세심해요 등) — 이모지 포함 추출
    const tags = [];
    // 이모지 유니코드 범위 (Emoticons, Symbols, Misc Symbols, Dingbats 등)
    const EMOJI_RE = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{1F000}-\u{1F02F}]/u;
    el.querySelectorAll('*').forEach(n => {
      if (n.children.length > 0) return; // leaf만
      const t = (n.textContent || '').trim();
      if (t.length < 3 || t.length > 30) return;
      if (!/요$|해요$|세심해요|자세해요|있어요$/.test(t)) return;
      // 이모지는 부모 텍스트/img alt에서 찾기
      let emoji = '';
      const parent = n.parentElement;
      if (parent) {
        const parentText = (parent.textContent || '').trim();
        const m = parentText.match(new RegExp('(' + EMOJI_RE.source + ')', 'u'));
        if (m) emoji = m[1];
        if (!emoji) {
          const emojiImg = parent.querySelector('img[alt]');
          if (emojiImg && EMOJI_RE.test(emojiImg.alt)) emoji = emojiImg.alt;
        }
      }
      const fullTag = emoji ? `${emoji} ${t}` : t;
      if (!tags.includes(fullTag)) tags.push(fullTag);
    });

    // 사진 URL
    const photos = [];
    el.querySelectorAll('img').forEach(img => {
      const src = img.src || img.getAttribute('data-src') || '';
      // 작은 썸네일/아이콘 제외 (프로필 이미지 등)
      if (!src || src.startsWith('data:')) return;
      if (img.width && img.width < 60) return;
      // 네이버 리뷰 이미지 패턴만 통과
      if (/phinf\.pstatic\.net|ldb-phinf\.pstatic\.net|blogfiles\.pstatic\.net|post-phinf\.pstatic\.net/.test(src)) {
        // 썸네일 쿼리 제거하여 원본 받기 시도
        const cleanUrl = src.replace(/\?type=[^&]+/, '').replace(/&w=\d+&h=\d+/, '');
        if (!photos.includes(cleanUrl)) photos.push(cleanUrl);
      }
    });

    // 영상 URL
    const videos = [];
    el.querySelectorAll('video').forEach(v => {
      const src = v.src || v.querySelector('source')?.src || '';
      if (src && !videos.includes(src)) videos.push(src);
    });
    // data-video 속성이 있는 요소도 체크
    el.querySelectorAll('[data-video-src], [data-videosrc]').forEach(n => {
      const src = n.getAttribute('data-video-src') || n.getAttribute('data-videosrc');
      if (src && !videos.includes(src)) videos.push(src);
    });

    const folder = `${String(idx+1).padStart(3,'0')}_${writeDate || 'nodate'}_${safeFolder(reviewer || 'unknown')}`;

    return {
      index: idx + 1,
      reviewer,
      write_date: writeDate,
      visit_date: visitDate,
      visit_time: visitTime,
      status,
      business,
      service,
      content,
      tags,
      photos,
      videos,
      folder,
    };
  }

  // ─────────────────────────────────────────────────────────────
  // 5. 전체 카드 파싱
  // ─────────────────────────────────────────────────────────────
  console.log('📝 리뷰 데이터 파싱 중...');
  const reviews = [];
  cards.forEach((card, i) => {
    try {
      reviews.push(extractCard(card, i));
    } catch (e) {
      console.error(`❌ 리뷰 #${i+1} 파싱 실패:`, e);
    }
  });
  console.log(`✅ 파싱 완료: ${reviews.length}건`);
  console.table(reviews.map(r => ({
    idx: r.index,
    name: r.reviewer,
    write: r.write_date,
    visit: r.visit_date,
    photos: r.photos.length,
    videos: r.videos.length,
    tags: r.tags.length,
  })));

  // ─────────────────────────────────────────────────────────────
  // 6. ZIP 생성 + 이미지/영상 다운로드
  // ─────────────────────────────────────────────────────────────
  const zip = new JSZip();
  const failed = [];
  const totalFiles = reviews.reduce((s, r) => s + r.photos.length + r.videos.length, 0);
  let doneFiles = 0;

  console.log(`📥 파일 다운로드 시작 (총 ${totalFiles}개)`);

  // 리뷰 메타데이터를 각 폴더의 review.md로 저장
  function buildReviewMd(r) {
    const lines = [];
    lines.push(`# ${r.reviewer || '(이름없음)'} 고객 리뷰`);
    lines.push('');
    lines.push('## 메타데이터');
    lines.push(`- **작성일**: ${r.write_date || '-'}`);
    lines.push(`- **방문일**: ${r.visit_date || '-'}${r.visit_time ? ' ' + r.visit_time : ''}`);
    if (r.status) lines.push(`- **상태**: ${r.status}`);
    if (r.business) lines.push(`- **업체**: ${r.business}`);
    if (r.service) lines.push(`- **서비스**: ${r.service}`);
    lines.push(`- **출처**: 네이버 스마트플레이스`);
    lines.push('');
    if (r.tags.length > 0) {
      lines.push('## 키워드');
      r.tags.forEach(t => lines.push(`- ${t}`));
      lines.push('');
    }
    lines.push('## 리뷰 본문');
    lines.push('');
    lines.push(r.content || '(본문 없음)');
    lines.push('');
    if (r.photos.length > 0) {
      lines.push('## 사진');
      r.photos.forEach((_, i) => lines.push(`- photo_${String(i+1).padStart(2,'0')}.jpg`));
      lines.push('');
    }
    if (r.videos.length > 0) {
      lines.push('## 영상');
      r.videos.forEach((_, i) => lines.push(`- video_${String(i+1).padStart(2,'0')}.mp4`));
      lines.push('');
    }
    return lines.join('\r\n');
  }

  for (const r of reviews) {
    const folder = zip.folder(r.folder);

    // review.md 저장 (본문+태그+메타데이터)
    folder.file('review.md', buildReviewMd(r));

    // 사진
    for (let pi = 0; pi < r.photos.length; pi++) {
      const url = r.photos[pi];
      const blob = await fetchBlob(url);
      doneFiles++;
      const fileName = `photo_${String(pi+1).padStart(2,'0')}`;
      if (!blob) {
        failed.push({ folder: r.folder, type: 'photo', fileName, url });
        continue;
      }
      const ext = guessExt(url, blob);
      folder.file(`${fileName}.${ext}`, blob);
      if (doneFiles % 20 === 0) {
        console.log(`📥 진행: ${doneFiles}/${totalFiles} (${Math.round(doneFiles/totalFiles*100)}%)`);
      }
    }

    // 영상
    for (let vi = 0; vi < r.videos.length; vi++) {
      const url = r.videos[vi];
      const blob = await fetchBlob(url);
      doneFiles++;
      const fileName = `video_${String(vi+1).padStart(2,'0')}`;
      if (!blob) {
        failed.push({ folder: r.folder, type: 'video', fileName, url });
        continue;
      }
      const ext = guessExt(url, blob);
      folder.file(`${fileName}.${ext}`, blob);
    }
  }

  // reviews.json + _failed.json
  zip.file('reviews.json', JSON.stringify(reviews, null, 2));
  if (failed.length > 0) {
    zip.file('_failed.json', JSON.stringify(failed, null, 2));
    console.warn(`⚠ 브라우저 CORS로 ${failed.length}건 다운로드 실패 — PowerShell 스크립트 자동 생성`);

    // URL 목록 TSV 파일 (Path\tURL) — 탭 구분으로 안전하게 저장
    const tsvLines = failed.map(f => `${f.folder}/${f.fileName}.jpg\t${f.url}`);
    zip.file('urls.tsv', tsvLines.join('\n'));

    // PowerShell 다운로드 스크립트 — 완전 ASCII (한글 없음, UTF-8 BOM 회피)
    // URL은 외부 TSV 파일에서 읽어옴 → 스크립트 본문에 & 문자 없음
    const ps1 = [
      '# MAMORU Naver Review Image Downloader',
      '# Reads URLs from urls.tsv (tab-separated: path\\tURL)',
      '',
      '$ErrorActionPreference = "Continue"',
      '$ProgressPreference = "SilentlyContinue"',
      '',
      '$root = Split-Path -Parent $MyInvocation.MyCommand.Definition',
      'Set-Location $root',
      '',
      '$lines = Get-Content -Path ".\\urls.tsv" -Encoding UTF8',
      '$total = $lines.Count',
      '$ok = 0',
      '$err = 0',
      '$i = 0',
      '',
      'Write-Host ""',
      'Write-Host "Downloading $total files..." -ForegroundColor Green',
      'Write-Host ""',
      '',
      'foreach ($line in $lines) {',
      '  $i++',
      '  if ([string]::IsNullOrWhiteSpace($line)) { continue }',
      '  $parts = $line -split "`t", 2',
      '  if ($parts.Count -lt 2) { continue }',
      '  $path = $parts[0]',
      '  $url  = $parts[1]',
      '  $dir  = Split-Path $path -Parent',
      '',
      '  try {',
      '    if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }',
      '    Invoke-WebRequest -Uri $url -OutFile $path -UseBasicParsing',
      '    $ok++',
      '    Write-Host ("  [OK]   [{0}/{1}] {2}" -f $i, $total, $path)',
      '  } catch {',
      '    $err++',
      '    Write-Host ("  [FAIL] [{0}/{1}] {2} - {3}" -f $i, $total, $path, $_.Exception.Message) -ForegroundColor Red',
      '  }',
      '}',
      '',
      'Write-Host ""',
      'Write-Host "Done!" -ForegroundColor Green',
      'Write-Host "  Success: $ok / Failed: $err"',
      'Write-Host ""',
      'Write-Host "Press Enter to exit..."',
      '$null = Read-Host',
      ''
    ].join('\r\n');

    // UTF-8 BOM 포함 (PowerShell 5.1 호환)
    zip.file('download_images.ps1', '\uFEFF' + ps1);

    // README
    zip.file('_README.txt',
      'MAMORU 네이버 리뷰 추출 결과\r\n' +
      '================================\r\n\r\n' +
      '⚠ 브라우저 CORS 정책으로 이미지는 직접 다운로드되지 않았습니다.\r\n' +
      '   URL은 reviews.json 과 urls.tsv 에 기록되어 있습니다.\r\n\r\n' +
      '📥 이미지 다운로드 방법:\r\n' +
      '   1. 이 ZIP을 원하는 폴더에 압축해제\r\n' +
      '   2. 폴더에서 Shift+우클릭 → "여기에서 PowerShell 창 열기"\r\n' +
      '   3. 다음 명령어 실행:\r\n' +
      '      powershell -ExecutionPolicy Bypass -File .\\download_images.ps1\r\n' +
      '   4. 완료되면 각 리뷰 폴더에 이미지가 저장됨\r\n\r\n' +
      '📂 파일 구조:\r\n' +
      '   reviews.json            — 전체 리뷰 메타데이터 (본문/태그/날짜)\r\n' +
      '   _failed.json            — 다운로드 실패 상세 (디버그용)\r\n' +
      '   urls.tsv                — 다운로드할 경로+URL 목록 (탭 구분)\r\n' +
      '   download_images.ps1     — PowerShell 다운로드 스크립트\r\n' +
      '   001_날짜_이름/          — 고객별 폴더\r\n' +
      '     ├─ review.md          — 본문+태그+메타데이터 (마크다운)\r\n' +
      '     ├─ photo_01.jpg       — 사진 (PS1 실행 후 들어감)\r\n' +
      '     └─ photo_02.jpg\r\n'
    );
  }

  // ─────────────────────────────────────────────────────────────
  // 7. ZIP 다운로드
  // ─────────────────────────────────────────────────────────────
  console.log('📦 ZIP 파일 생성 중...');
  const zipBlob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  const today = new Date().toISOString().slice(0, 10);
  const fileName = `naver_reviews_${today}.zip`;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(zipBlob);
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  console.log(`%c✅ 완료! ${fileName} 다운로드됨`, 'font-size:16px;font-weight:bold;color:#03C75A');
  console.log(`   총 리뷰: ${reviews.length}건`);
  console.log(`   브라우저 다운로드 성공: ${doneFiles - failed.length}/${totalFiles}`);
  if (failed.length > 0) {
    console.log(`%c   ⚠ CORS로 실패 ${failed.length}건 → download_images.ps1 실행하세요`, 'color:#F59E0B');
    console.log(`%c   💡 ZIP 압축해제 → download_images.ps1 우클릭 → PowerShell로 실행`, 'color:#3B82F6');
  }

  // 전역 window.mamoruReviews로 노출 (추가 작업용)
  window.mamoruReviews = reviews;
  console.log('💡 window.mamoruReviews 에 전체 데이터 저장됨 (Console에서 추가 확인 가능)');
})();
