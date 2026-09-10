#!/usr/bin/env python3
"""
마모루 유튜브 검색어 자동 추출 v2 (YouTube Data API v3 + 자동완성 → Notion)

무엇을 하나
  1) 씨앗 검색어(가위 직접 / 미용사 실무 / 고민·증상)를 유튜브에 검색해 상위 10개 영상을 가져온다
  2) 영상 조회수·게시일 + 채널 구독자 수로 '기회 점수(0~100)'를 매긴다
       💎 빈자리 (60점↑) : 상위가 소형 채널·오래된 영상 → 최신 영상으로 들어가면 자리 있음
       🟡 보통  (36~59)  : 썸네일·첫 30초로 승부
       🔥 대형 판(35점↓) : 대형 채널 최신 영상이 꽉 잡음 → 장기전·시리즈용
  3) 유튜브 자동완성으로 "미용사가 실제로 치는 표현"을 모은다 (🔎)
  4) 실무·고민 검색어 상위에 자주 뜨는 '미용 채널'을 자동 선정해 최근 3개월 인기 제목을 기록한다 (👀)
  5) 결과를 표 + 설명 + 조합 가이드로 Notion '🔍 주간 검색어 로그' DB에 기록한다 (채널=유튜브)

비밀값은 전부 환경변수(=GitHub Secrets)에서 읽는다. 코드에 키 없음.
하루 할당량 10,000유닛 중 약 5,000유닛 사용 (search.list 1회 = 100유닛).
"""
import os
import sys
import time
import datetime
import statistics
from collections import Counter, defaultdict

import requests

# ── 비밀값 ──
YOUTUBE_API_KEY = os.environ["YOUTUBE_API_KEY"]
NOTION_TOKEN = os.environ["NOTION_TOKEN"]
NOTION_DB_ID = os.environ.get("NOTION_DB_ID", "4c94f972-7936-4396-8586-516eb6a10dea")

# ── 씨앗 검색어 3층. 여기만 고치면 조사 범위가 바뀜 ──
SEEDS = {
    "① 가위 직접": [
        "미용가위", "틴닝가위", "숱가위", "커트가위", "미용가위 추천",
        "가위 연마", "가위 수리", "일본가위",
    ],
    "② 실무 기법": [
        "슬라이싱", "포인트컷", "블런트컷", "틴닝 기법", "커트 연습",
        "미용가위 잡는법", "위그 커트 연습", "앞머리 커트",
    ],
    "③ 고민·증상": [
        "가위 안들때", "가위 소리", "가위 뻑뻑", "손목 아픈 미용사",
        "신입 미용사", "미용사 취업", "미용가위 관리",
    ],
}
LONGTAIL_TO_SCORE = 10          # 자동완성 롱테일 중 상위 영상까지 조사할 개수 (1개 = 100유닛)
WATCH_CANDIDATES = 15           # 👀 후보 채널 수 (미용 채널인지 확인 후 아래 수만 남김)
WATCH_CHANNELS = 10
WATCH_VIDEOS_PER_CHANNEL = 2

# 미용가위와 무관한 영상·채널·표현 걸러내기 (제목·채널명에 이 단어가 있으면 제외)
NOISE_TOKENS = [
    "애견", "강아지", "펫", "반려", "고양이", "코털",
    "골다공증", "효능", "관절", "염증", "혈압", "당뇨", "건강식",
    "대기업", "연봉", "서울대", "공무원", "면접",
    "주방", "요리", "원예", "정원", "종이", "손톱", "네일", "재봉", "공예", "게임", "먹방",
]
# 미용 채널 판별 (👀): 채널명·인기 제목에 이 중 하나는 있어야 함
HAIR_TOKENS = ["미용", "커트", "헤어", "가위", "펌", "디자이너", "미용실", "틴닝", "숱", "시저", "hair", "barber", "이용"]

YT = "https://www.googleapis.com/youtube/v3"
_units_used = 0


# ───────────────────────── YouTube API ─────────────────────────
def yt_get(endpoint: str, units: int, **params):
    global _units_used
    params["key"] = YOUTUBE_API_KEY
    r = requests.get(f"{YT}/{endpoint}", params=params, timeout=30)
    _units_used += units
    if r.status_code >= 300:
        raise RuntimeError(f"{endpoint} {r.status_code}: {r.text[:200]}")
    return r.json()


def is_noise(text: str) -> bool:
    low = (text or "").lower()
    return any(t in low for t in NOISE_TOKENS)


def is_hair(text: str) -> bool:
    low = (text or "").lower()
    return any(t in low for t in HAIR_TOKENS)


def search_top(query: str, n: int = 10, **extra):
    """검색어의 상위 n개 영상 (유튜브 검색 결과 순서). 100유닛. 노이즈 영상은 뺀다."""
    data = yt_get("search", 100, part="snippet", q=query, type="video", maxResults=n,
                  regionCode="KR", relevanceLanguage="ko", **extra)
    out = []
    for it in data.get("items", []):
        vid = it.get("id", {}).get("videoId")
        sn = it.get("snippet", {})
        if not vid:
            continue
        if is_noise(sn.get("title", "")) or is_noise(sn.get("channelTitle", "")):
            continue
        out.append({"videoId": vid, "channelId": sn.get("channelId"),
                    "channelTitle": sn.get("channelTitle", ""), "title": sn.get("title", ""),
                    "publishedAt": sn.get("publishedAt", "")})
    return out


def chunks(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def fetch_video_stats(video_ids):
    stats = {}
    for batch in chunks(list(dict.fromkeys(video_ids)), 50):
        try:
            data = yt_get("videos", 1, part="statistics", id=",".join(batch))
            for it in data.get("items", []):
                stats[it["id"]] = int(it.get("statistics", {}).get("viewCount", 0))
        except Exception as e:  # noqa: BLE001
            print(f"[warn] videos.list 실패: {e}", file=sys.stderr)
    return stats


def fetch_channel_subs(channel_ids):
    subs = {}
    for batch in chunks(list(dict.fromkeys(c for c in channel_ids if c)), 50):
        try:
            data = yt_get("channels", 1, part="statistics", id=",".join(batch))
            for it in data.get("items", []):
                st = it.get("statistics", {})
                subs[it["id"]] = None if st.get("hiddenSubscriberCount") else int(st.get("subscriberCount", 0))
        except Exception as e:  # noqa: BLE001
            print(f"[warn] channels.list 실패: {e}", file=sys.stderr)
    return subs


# ───────────────────────── 자동완성 (무료, 할당량 없음) ─────────────────────────
def autocomplete(query: str):
    try:
        r = requests.get("https://suggestqueries.google.com/complete/search",
                         params={"client": "firefox", "ds": "yt", "hl": "ko", "gl": "kr", "q": query},
                         timeout=15)
        r.raise_for_status()
        return [s for s in r.json()[1] if isinstance(s, str)]
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 자동완성 실패 '{query}': {e}", file=sys.stderr)
        return []


# ───────────────────────── 점수 계산 ─────────────────────────
def years_ago(iso: str) -> float:
    try:
        d = datetime.datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return (datetime.datetime.now(datetime.timezone.utc) - d).days / 365
    except Exception:  # noqa: BLE001
        return 0.0


def fmt_num(n):
    if n is None:
        return "비공개"
    if n >= 10000:
        return f"{n / 10000:.1f}만".replace(".0만", "만")
    return f"{n:,}"


def score_keyword(kw: str, videos, stats, subs):
    """기회 점수 0~100 = 오래됨 45 + 소형채널 35 + 수요 20"""
    if len(videos) < 3:
        return None
    sub_list = [subs.get(v["channelId"]) for v in videos]
    sub_list = [s for s in sub_list if s is not None]
    views = [stats.get(v["videoId"], 0) for v in videos]
    n = len(videos)
    old = sum(1 for v in videos if years_ago(v["publishedAt"]) >= 2)
    med_subs = int(statistics.median(sub_list)) if sub_list else 0
    avg_views = int(statistics.mean(views)) if views else 0

    old_ratio = old / n
    smallness = 1 - min(med_subs, 200_000) / 200_000
    demand = min(avg_views, 500_000) / 500_000
    score = round(45 * old_ratio + 35 * smallness + 20 * demand)
    if avg_views < 3_000:
        grade = "mid"          # 수요가 너무 작으면 빈자리라도 보통으로
    elif score >= 60:
        grade = "gem"
    elif score <= 35:
        grade = "hard"
    else:
        grade = "mid"
    newest = max(videos, key=lambda v: v["publishedAt"])
    return {"kw": kw, "grade": grade, "score": score, "med_subs": med_subs, "avg_views": avg_views,
            "old": old, "n": n, "newest_channel": newest["channelTitle"]}


# ───────────────────────── 수집 ─────────────────────────
def collect():
    all_seeds = [(tier, kw) for tier, kws in SEEDS.items() for kw in kws]

    # 1) 자동완성
    longtail = Counter()
    for _, kw in all_seeds:
        for s in autocomplete(kw):
            if s != kw and not is_noise(s):
                longtail[s] += 1
        time.sleep(0.2)
    longtail_sorted = [s for s, _ in longtail.most_common(60)]

    # 2) 상위 영상 조회
    to_score = all_seeds + [("🔎 롱테일", kw) for kw in longtail_sorted[:LONGTAIL_TO_SCORE]]
    results = {}
    for tier, kw in to_score:
        try:
            results[kw] = {"tier": tier, "videos": search_top(kw)}
        except Exception as e:  # noqa: BLE001
            print(f"[warn] 검색 실패 '{kw}': {e}", file=sys.stderr)
            results[kw] = {"tier": tier, "videos": []}
        time.sleep(0.2)

    all_videos = [v for r in results.values() for v in r["videos"]]
    stats = fetch_video_stats([v["videoId"] for v in all_videos])
    subs = fetch_channel_subs([v["channelId"] for v in all_videos])

    scored = []
    for kw, r in results.items():
        s = score_keyword(kw, r["videos"], stats, subs)
        if s:
            s["tier"] = r["tier"]
            scored.append(s)

    # 3) 👀 미용사가 보는 채널: ②③층 상위에 '서로 다른 검색어 2개 이상'에서 뜬 채널
    kw_hits = defaultdict(set)
    names = {}
    for kw, r in results.items():
        if r["tier"].startswith("②") or r["tier"].startswith("③"):
            for v in r["videos"]:
                cid = v["channelId"]
                if cid and (subs.get(cid) or 0) >= 1000:
                    kw_hits[cid].add(kw)
                    names[cid] = v["channelTitle"]
    candidates = sorted(kw_hits.items(), key=lambda x: -len(x[1]))
    candidates = [(cid, kws) for cid, kws in candidates if len(kws) >= 2][:WATCH_CANDIDATES]

    watch = []
    since = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=90)).strftime("%Y-%m-%dT%H:%M:%SZ")
    for cid, kws in candidates:
        if len(watch) >= WATCH_CHANNELS:
            break
        try:
            data = yt_get("search", 100, part="snippet", channelId=cid, type="video", order="viewCount",
                          publishedAfter=since, maxResults=WATCH_VIDEOS_PER_CHANNEL)
            vids = [{"videoId": it["id"]["videoId"], "title": it["snippet"]["title"]} for it in data.get("items", [])
                    if it.get("id", {}).get("videoId")]
        except Exception as e:  # noqa: BLE001
            print(f"[warn] 채널 인기영상 실패 {names.get(cid)}: {e}", file=sys.stderr)
            vids = []
        # 미용 채널인지 확인: 채널명 또는 인기 제목에 미용 단어가 있고, 노이즈 없음
        texts = [names.get(cid, "")] + [v["title"] for v in vids]
        if not any(is_hair(t) for t in texts) or any(is_noise(t) for t in texts):
            continue
        vstats = fetch_video_stats([v["videoId"] for v in vids]) if vids else {}
        watch.append({"channel": names.get(cid, cid), "subs": subs.get(cid), "hits": len(kws), "kws": sorted(kws),
                      "videos": [{"title": v["title"], "views": vstats.get(v["videoId"], 0)} for v in vids]})
        time.sleep(0.2)

    return scored, longtail_sorted, watch


# ───────────────────────── Notion 블록 ─────────────────────────
def week_label(d: datetime.date) -> str:
    nth = (d.day - 1) // 7 + 1
    return f"{d.month}월 {nth}주차 ({d:%m/%d}) · 유튜브"


def _rt(text, bold=False):
    return [{"type": "text", "text": {"content": str(text)[:1900]}, "annotations": {"bold": bold}}]


def _heading(text):
    return {"object": "block", "type": "heading_3", "heading_3": {"rich_text": _rt(text)}}


def _para(text):
    return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": _rt(text)}}


def _bullet(text):
    return {"object": "block", "type": "bulleted_list_item", "bulleted_list_item": {"rich_text": _rt(text)}}


def _todo(text):
    return {"object": "block", "type": "to_do", "to_do": {"rich_text": _rt(text), "checked": False}}


def _callout(text, emoji="✍️", color="gray_background"):
    return {"object": "block", "type": "callout",
            "callout": {"icon": {"emoji": emoji}, "color": color, "rich_text": _rt(text)}}


def _toggle(title, children):
    return {"object": "block", "type": "toggle", "toggle": {"rich_text": _rt(title), "children": children[:100]}}


def _table(header, rows):
    """Notion 표 블록. header: 열 이름 리스트, rows: 셀 문자열 리스트의 리스트"""
    def row(cells, bold=False):
        return {"object": "block", "type": "table_row",
                "table_row": {"cells": [_rt(c, bold) for c in cells]}}
    return {"object": "block", "type": "table",
            "table": {"table_width": len(header), "has_column_header": True, "has_row_header": False,
                      "children": [row(header, True)] + [row(r) for r in rows[:98]]}}


GRADE_ICON = {"gem": "💎", "mid": "🟡", "hard": "🔥"}


def kw_row(s):
    return [s["tier"][:1], s["kw"], str(s["score"]), fmt_num(s["med_subs"]), fmt_num(s["avg_views"]),
            f"{s['old']}/{s['n']}", s["newest_channel"]]


def build_children(scored, longtail, watch):
    header = ["층", "검색어", "점수", "구독자 중앙값", "평균 조회", "2년↑", "최신 상위 채널"]
    by_grade = {g: sorted([s for s in scored if s["grade"] == g], key=lambda s: -s["score"]) for g in GRADE_ICON}

    ch = []
    # ── 읽는 법 ──
    ch.append(_callout(
        "숫자 읽는 법 — "
        "층: ① 가위 직접 / ② 실무 기법 / ③ 고민·증상 / 🔎 자동완성에서 나온 롱테일.  "
        "점수(0~100): 기회 점수. 오래된 영상 비율 45 + 상위 채널이 작을수록 35 + 수요(조회) 20. 60점↑ 💎 / 36~59 🟡 / 35↓ 🔥.  "
        "구독자 중앙값: 상위 10개 영상을 올린 채널들의 구독자 '가운데 값' — 작을수록 소형 채널도 상위에 있다 = 나도 올라갈 수 있다.  "
        "평균 조회: 상위 10개 평균 조회수 — 수요 크기.  "
        "2년↑: 상위 10개 중 2년 넘은 영상 수 — 많을수록 최신 영상이 비집고 들어갈 자리.  "
        "최신 상위: 상위 10개 중 가장 최근 영상의 채널 — 지금 그 자리를 잡고 있는 경쟁자.",
        "📐"))

    ch.append(_heading("💎 가위 관점 빈자리 (다음 영상 후보)"))
    ch.append(_table(header, [kw_row(s) for s in by_grade["gem"]]) if by_grade["gem"] else _bullet("해당 없음"))
    ch.append(_heading("🟡 보통 (썸네일·첫 30초로 승부)"))
    ch.append(_table(header, [kw_row(s) for s in by_grade["mid"]]) if by_grade["mid"] else _bullet("해당 없음"))
    ch.append(_heading("🔥 대형 채널 판 (장기전 · 시리즈용)"))
    ch.append(_table(header, [kw_row(s) for s in by_grade["hard"]]) if by_grade["hard"] else _bullet("해당 없음"))

    ch.append(_heading("🔎 미용사가 실제로 치는 표현 (자동완성 · 제목·설명에 그대로 쓸 말)"))
    ch += [_bullet(" · ".join(longtail[i:i + 6])) for i in range(0, min(len(longtail), 48), 6)] or [_bullet("해당 없음")]

    ch.append(_heading("👀 미용사가 요즘 보는 것 (자동 선정 미용 채널 · 최근 3개월 인기 제목 — 따라 하기 ❌, 가위 관점으로 비틀기)"))
    if watch:
        rows = []
        for w in watch:
            titles = " / ".join(f"{v['title'][:45]} ({fmt_num(v['views'])})" for v in w["videos"]) or "-"
            rows.append([w["channel"], fmt_num(w["subs"]), ", ".join(w["kws"][:3]), titles])
        ch.append(_table(["채널", "구독자", "어떤 검색어 상위", "최근 3개월 인기 제목 (조회)"], rows))
    else:
        ch.append(_bullet("해당 없음"))

    # ── 조합 가이드 ──
    top = by_grade["gem"][:3] or by_grade["mid"][:3]
    ex_kw = top[0]["kw"] if top else "미용가위 종류"
    ex_lt = next((l for l in longtail if ex_kw.split()[0] in l and l != ex_kw), "고르는 법")
    ch.append(_heading("🧩 조합 가이드 — 제목·썸네일 만들 때"))
    ch.append(_table(
        ["자리", "어디서 고르나", "규칙", "이번 주 예시"],
        [
            ["제목 앞쪽", "💎 표에서 1개", "검색어 그대로, 앞쪽에", ex_kw],
            ["제목 뒷부분 · 설명 첫 줄", "🔎 표현에서 같은 뿌리 1개", "미용사가 치는 말투 그대로", ex_lt],
            ["썸네일 한 줄 (8자 이내)", "👀 인기 제목의 '형식'만", "질문 또는 반전, 낚시 ❌, 실명 ❌", "\"버리라던 가위입니다\" 계열"],
            ["형식", "👀 표", "쇼츠·ASMR·전/후가 많으면 그 형식으로", "복원 전/후 15~30초 + 롱폼 링크"],
        ]))
    ch.append(_para(f"예) 제목: \"{ex_kw}, {ex_lt} — 브랜드보다 먼저 볼 것\"  ·  썸네일: \"비싼 가위 = 잘 잘린다?\"  ·  설명 첫 줄에 {ex_kw} 1회 + 표준 링크 1개."))
    ch.append(_callout(
        "Claude에게 이렇게 요청: \"이번 주 유튜브 로그 💎 상위 3개로 제목 후보 5개 + 썸네일 문구 3개 뽑아줘 (브랜드 톤)\". "
        "쓴 검색어는 아래 토글의 ☑ 체크로 사용완료 표시.", "✍️"))
    all_sorted = sorted(scored, key=lambda s: -s["score"])
    ch.append(_toggle("☑ 사용완료 체크용 (전체 검색어)",
                      [_todo(f"{GRADE_ICON[s['grade']]} [{s['tier'][:1]}] {s['kw']} ({s['score']}점)") for s in all_sorted]))
    return ch[:100]


def create_notion_entry(scored, longtail, watch):
    today = datetime.date.today()
    gems = [s["kw"] for s in sorted(scored, key=lambda s: -s["score"]) if s["grade"] == "gem"][:7]
    payload = {
        "parent": {"database_id": NOTION_DB_ID},
        "properties": {
            "주차": {"title": [{"text": {"content": week_label(today)}}]},
            "날짜": {"date": {"start": today.isoformat()}},
            "검색어": {"rich_text": [{"text": {"content": ", ".join(gems) or "-"}}]},
            "채널": {"multi_select": [{"name": "유튜브"}]},
            "반영": {"checkbox": False},
            "메모": {"rich_text": [{"text": {"content": f"GitHub Actions 자동 기록 (YouTube Data API · 자동완성 · 약 {_units_used:,}유닛 사용)"}}]},
        },
        "children": build_children(scored, longtail, watch),
    }
    r = requests.post("https://api.notion.com/v1/pages",
                      headers={"Authorization": f"Bearer {NOTION_TOKEN}", "Notion-Version": "2022-06-28",
                               "Content-Type": "application/json"},
                      json=payload, timeout=60)
    if r.status_code >= 300:
        print(f"[error] Notion 기록 실패 {r.status_code}: {r.text}", file=sys.stderr)
        r.raise_for_status()
    print(f"[ok] Notion 기록 완료 — {week_label(today)} · 검색어 {len(scored)}개 · 롱테일 {len(longtail)}개 · 채널 {len(watch)}개 · {_units_used:,}유닛")


def main():
    scored, longtail, watch = collect()
    if not scored:
        print("[error] 수집된 결과가 없습니다. YOUTUBE_API_KEY 또는 API 사용 설정을 확인하세요.", file=sys.stderr)
        sys.exit(1)
    create_notion_entry(scored, longtail, watch)


if __name__ == "__main__":
    main()
