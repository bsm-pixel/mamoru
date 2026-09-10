#!/usr/bin/env python3
"""
마모루 유튜브 검색어 자동 추출 (YouTube Data API v3 + 자동완성 → Notion)

무엇을 하나
  1) 씨앗 검색어(가위 직접 / 미용사 실무 / 고민·증상)를 유튜브에 검색해 상위 10개 영상을 가져온다
  2) 그 영상들의 조회수·게시일 + 채널 구독자 수를 보고 "빈자리"인지 판단한다
       💎 가위 관점 빈자리  : 상위가 소형 채널이거나 2년 넘은 영상 위주 → 다음 영상 후보
       🔥 대형 채널 판      : 상위가 대형 채널 최신 영상으로 꽉 참 → 장기전·시리즈용
  3) 유튜브 자동완성으로 "미용사가 실제로 치는 표현"을 모은다 (🔎)
  4) 실무·고민 검색어 상위에 자주 뜨는 채널 10개를 자동 선정해 최근 3개월 인기 제목을 기록한다 (👀)
  5) 결과를 Notion '🔍 주간 검색어 로그' DB에 새 항목(채널=유튜브)으로 기록한다

비밀값은 전부 환경변수(=GitHub Secrets)에서 읽는다. 코드에 키 없음.
하루 할당량 10,000유닛 중 약 4,500유닛 사용 (search.list 1회 = 100유닛).
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
        "가위 잡는법", "위그 커트", "앞머리 커트",
    ],
    "③ 고민·증상": [
        "가위 안들때", "가위 소리", "가위 뻑뻑", "손목 아픈 미용사",
        "신입 미용사", "미용 취업", "가위 관리",
    ],
}
# 자동완성에서 나온 롱테일 중 몇 개까지 추가로 상위 영상까지 조사할지 (1개 = 100유닛)
LONGTAIL_TO_SCORE = 10
# 👀 미용사가 보는 채널: 몇 개 채널 / 채널당 몇 개 영상
WATCH_CHANNELS = 10
WATCH_VIDEOS_PER_CHANNEL = 2

# 관련성 필터 (자동완성 결과용). 미용가위와 무관한 걸 걸러냄
EXCLUDE_TOKENS = [
    "코털", "강아지", "애견", "반려", "펫", "고양이", "주방", "요리", "원예", "정원",
    "종이", "손톱", "네일", "재봉", "공예", "게임", "노래", "asmr 먹방",
]

YT = "https://www.googleapis.com/youtube/v3"
_units_used = 0


# ───────────────────────── YouTube API ─────────────────────────
def yt_get(endpoint: str, units: int, **params):
    """YouTube Data API 호출. 실패해도 봇 전체가 죽지 않게 예외는 호출부에서 처리."""
    global _units_used
    params["key"] = YOUTUBE_API_KEY
    r = requests.get(f"{YT}/{endpoint}", params=params, timeout=30)
    _units_used += units
    if r.status_code >= 300:
        raise RuntimeError(f"{endpoint} {r.status_code}: {r.text[:200]}")
    return r.json()


def search_top(query: str, n: int = 10, **extra):
    """검색어의 상위 n개 영상 (relevance = 유튜브 검색 결과 순서). 100유닛."""
    data = yt_get("search", 100, part="snippet", q=query, type="video", maxResults=n,
                  regionCode="KR", relevanceLanguage="ko", **extra)
    out = []
    for it in data.get("items", []):
        vid = it.get("id", {}).get("videoId")
        sn = it.get("snippet", {})
        if vid:
            out.append({"videoId": vid, "channelId": sn.get("channelId"),
                        "channelTitle": sn.get("channelTitle", ""), "title": sn.get("title", ""),
                        "publishedAt": sn.get("publishedAt", "")})
    return out


def chunks(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def fetch_video_stats(video_ids):
    """조회수. 50개당 1유닛."""
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
    """구독자 수. 50개당 1유닛."""
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
    """유튜브 검색창 자동완성 = 사람들이 실제로 치는 표현."""
    try:
        r = requests.get("https://suggestqueries.google.com/complete/search",
                         params={"client": "firefox", "ds": "yt", "hl": "ko", "gl": "kr", "q": query},
                         timeout=15)
        r.raise_for_status()
        data = r.json()
        return [s for s in data[1] if isinstance(s, str)]
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 자동완성 실패 '{query}': {e}", file=sys.stderr)
        return []


def is_relevant(text: str) -> bool:
    low = text.lower()
    return not any(t in low for t in EXCLUDE_TOKENS)


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
    """상위 영상 묶음으로 '빈자리' 여부 판단."""
    if not videos:
        return None
    sub_list = [subs.get(v["channelId"]) for v in videos]
    sub_list = [s for s in sub_list if s is not None]
    views = [stats.get(v["videoId"], 0) for v in videos]
    old = sum(1 for v in videos if years_ago(v["publishedAt"]) >= 2)
    med_subs = int(statistics.median(sub_list)) if sub_list else 0
    avg_views = int(statistics.mean(views)) if views else 0
    n = len(videos)
    # 판정
    if med_subs < 50_000 or old >= n * 0.5:
        grade = "gem"     # 💎 빈자리
    elif med_subs >= 200_000 and old <= n * 0.2:
        grade = "hard"    # 🔥 대형 채널 판
    else:
        grade = "mid"     # 🟡 보통
    newest = max(videos, key=lambda v: v["publishedAt"])
    return {"kw": kw, "grade": grade, "med_subs": med_subs, "avg_views": avg_views,
            "old": old, "n": n, "newest_channel": newest["channelTitle"]}


def line(r):
    return (f"{r['kw']}  —  상위{r['n']} 구독자 중앙값 {fmt_num(r['med_subs'])} · 평균 조회 {fmt_num(r['avg_views'])}"
            f" · 2년↑ 영상 {r['old']}/{r['n']} · 최신 상위: {r['newest_channel']}")


# ───────────────────────── 수집 ─────────────────────────
def collect():
    all_seeds = [(tier, kw) for tier, kws in SEEDS.items() for kw in kws]

    # 1) 자동완성 — 씨앗마다 실제 치는 표현 수집
    longtail = Counter()
    for _, kw in all_seeds:
        for s in autocomplete(kw):
            if s != kw and is_relevant(s):
                longtail[s] += 1
        time.sleep(0.2)
    longtail_sorted = [s for s, _ in longtail.most_common(60)]

    # 2) 상위 영상 조회: 씨앗 전부 + 롱테일 상위 N개
    to_score = [(tier, kw) for tier, kw in all_seeds] + [("🔎 롱테일", kw) for kw in longtail_sorted[:LONGTAIL_TO_SCORE]]
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

    # 3) 👀 미용사가 보는 채널: ②③층 상위에 자주 뜨는 채널 (구독자 1천 이상만)
    freq = Counter()
    names = {}
    for kw, r in results.items():
        if r["tier"].startswith("②") or r["tier"].startswith("③"):
            for v in r["videos"]:
                cid = v["channelId"]
                if cid and (subs.get(cid) or 0) >= 1000:
                    freq[cid] += 1
                    names[cid] = v["channelTitle"]
    watch = []
    since = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=90)).strftime("%Y-%m-%dT%H:%M:%SZ")
    for cid, _ in freq.most_common(WATCH_CHANNELS):
        try:
            data = yt_get("search", 100, part="snippet", channelId=cid, type="video", order="viewCount",
                          publishedAfter=since, maxResults=WATCH_VIDEOS_PER_CHANNEL)
            vids = [{"videoId": it["id"]["videoId"], "title": it["snippet"]["title"]} for it in data.get("items", [])
                    if it.get("id", {}).get("videoId")]
        except Exception as e:  # noqa: BLE001
            print(f"[warn] 채널 인기영상 실패 {names.get(cid)}: {e}", file=sys.stderr)
            vids = []
        vstats = fetch_video_stats([v["videoId"] for v in vids]) if vids else {}
        watch.append({"channel": names.get(cid, cid), "subs": subs.get(cid), "hits": freq[cid],
                      "videos": [{"title": v["title"], "views": vstats.get(v["videoId"], 0)} for v in vids]})
        time.sleep(0.2)

    return scored, longtail_sorted, watch


# ───────────────────────── Notion ─────────────────────────
def week_label(d: datetime.date) -> str:
    nth = (d.day - 1) // 7 + 1
    return f"{d.month}월 {nth}주차 ({d:%m/%d}) · 유튜브"


def _rt(text):
    return [{"type": "text", "text": {"content": text[:1900]}}]


def _heading(text):
    return {"object": "block", "type": "heading_3", "heading_3": {"rich_text": _rt(text)}}


def _bullet(text):
    return {"object": "block", "type": "bulleted_list_item", "bulleted_list_item": {"rich_text": _rt(text)}}


def _todo(text):
    return {"object": "block", "type": "to_do", "to_do": {"rich_text": _rt(text), "checked": False}}


def _callout(text, emoji="✍️"):
    return {"object": "block", "type": "callout", "callout": {"icon": {"emoji": emoji}, "rich_text": _rt(text)}}


def build_children(scored, longtail, watch):
    gems = sorted([s for s in scored if s["grade"] == "gem"], key=lambda s: -s["avg_views"])
    hard = sorted([s for s in scored if s["grade"] == "hard"], key=lambda s: -s["avg_views"])
    mid = sorted([s for s in scored if s["grade"] == "mid"], key=lambda s: -s["avg_views"])

    ch = [_callout("판정 기준 — 💎: 상위10 채널 구독자 중앙값 5만 미만이거나 2년 넘은 영상이 절반 이상 (= 소형 채널도 올라갈 자리). "
                   "🔥: 상위가 구독자 20만 이상 대형 채널 최신 영상으로 꽉 참 (= 장기전). 🟡: 그 사이.", "📐")]
    ch.append(_heading("💎 가위 관점 빈자리 (다음 영상 후보)"))
    ch += [_todo(f"[{s['tier'][:1]}] {line(s)}") for s in gems] or [_bullet("해당 없음")]
    ch.append(_heading("🟡 보통 (썸네일·첫 30초로 승부)"))
    ch += [_todo(f"[{s['tier'][:1]}] {line(s)}") for s in mid[:12]] or [_bullet("해당 없음")]
    ch.append(_heading("🔥 수요 크지만 대형 채널 판 (장기전 · 시리즈용)"))
    ch += [_todo(f"[{s['tier'][:1]}] {line(s)}") for s in hard] or [_bullet("해당 없음")]

    ch.append(_heading("🔎 미용사가 실제로 치는 표현 (자동완성 · 제목·설명에 그대로 쓸 말)"))
    ch += [_bullet(" · ".join(longtail[i:i + 6])) for i in range(0, min(len(longtail), 48), 6)] or [_bullet("해당 없음")]

    ch.append(_heading("👀 미용사가 요즘 보는 것 (자동 선정 채널 · 최근 3개월 인기 제목 — 따라 하기 ❌, 가위 관점으로 비틀기)"))
    for w in watch:
        ch.append(_bullet(f"{w['channel']} (구독 {fmt_num(w['subs'])} · 실무 검색어 상위 {w['hits']}회)"))
        for v in w["videos"]:
            ch.append(_bullet(f"　└ {v['title']} — 조회 {fmt_num(v['views'])}"))
    if not watch:
        ch.append(_bullet("해당 없음"))

    ch.append(_callout("쓴 검색어는 ☑ 체크해 '사용완료'. 제목은 Claude에 요청: \"이 검색어로 유튜브 제목·썸네일 문구 뽑아줘\" (💎부터, [①②③] = 씨앗 층, 브랜드 톤)."))
    return ch[:100]  # Notion 페이지 생성 시 children 최대 100블록


def create_notion_entry(scored, longtail, watch):
    today = datetime.date.today()
    gems = [s["kw"] for s in scored if s["grade"] == "gem"][:7]
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
                      json=payload, timeout=30)
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
