#!/usr/bin/env python3
"""
마모루 주간 검색어 자동 추출
- 네이버 검색광고 API(키워드도구/RelKwdStat)로 씨앗 키워드의 연관검색어 + 월간검색량 수집
- 상위 키워드를 Notion '🔍 주간 검색어 로그' DB에 새 항목으로 기록
매주 GitHub Actions 크론으로 실행. 비밀값은 전부 환경변수(=GitHub Secrets)에서 읽는다.
"""
import os
import sys
import time
import hmac
import base64
import hashlib
import datetime
import requests

# ── 비밀값(=GitHub Secrets / 환경변수) ──
NAVER_API_KEY = os.environ["NAVER_API_KEY"]          # 액세스라이선스
NAVER_SECRET_KEY = os.environ["NAVER_SECRET_KEY"]    # 비밀키
NAVER_CUSTOMER_ID = os.environ["NAVER_CUSTOMER_ID"]  # CUSTOMER_ID
NOTION_TOKEN = os.environ["NOTION_TOKEN"]            # 노션 통합 토큰
NOTION_DB_ID = os.environ.get("NOTION_DB_ID", "4c94f972-7936-4396-8586-516eb6a10dea")

# ── 씨앗 키워드(미용사 대상 전문어 위주). 여기만 고치면 조사 범위가 바뀜 ──
SEED_KEYWORDS = [
    # 제품·용어 (미용 전문어)
    "미용가위", "틴닝가위", "커트가위", "장가위", "미용가위추천",
    # 관리·수리·연마 (마모루 강점)
    "가위수리", "미용가위연마", "미용가위관리",
    # 미용 특화 기술·입문 (일반 '가위'는 노이즈라 제외)
    "미용가위입문", "숱치기",
]
TOP_N = 15  # Notion에 기록할 상위 키워드 수

# 관련성 필터: 연관검색어 중 아래 토큰을 포함한 것만 남긴다
# (검색량만으로 정렬하면 롤·빗·바리깡 같은 일반 소품이 '가위' 전문어를 밀어내므로)
RELEVANT_TOKENS = ["가위", "시저스", "틴닝", "숱", "연마", "scissor"]
# 미용(헤어) 가위 아닌 것 배제 (반려동물·주방·사무·재봉·공예 등)
EXCLUDE_TOKENS = [
    "코털", "강아지", "애견", "반려", "펫", "고양이",
    "주방", "요리", "고기", "가지", "원예", "정원", "꽃",
    "전동", "핑킹", "사무", "안전", "프린텍", "재단", "종이", "손톱", "네일",
]

NAVER_BASE = "https://api.searchad.naver.com"


def is_relevant(kw: str) -> bool:
    low = kw.lower()
    if any(t in low for t in EXCLUDE_TOKENS):
        return False
    return any(t in low for t in RELEVANT_TOKENS)


def naver_signature(timestamp: str, method: str, path: str) -> str:
    msg = f"{timestamp}.{method}.{path}"
    digest = hmac.new(NAVER_SECRET_KEY.encode("utf-8"), msg.encode("utf-8"), hashlib.sha256).digest()
    return base64.b64encode(digest).decode("utf-8")


def fetch_keywords(hint_keywords):
    """RelKwdStat 호출 → keywordList 반환"""
    path = "/keywordstool"
    ts = str(int(time.time() * 1000))
    headers = {
        "X-Timestamp": ts,
        "X-API-KEY": NAVER_API_KEY,
        "X-Customer": str(NAVER_CUSTOMER_ID),
        "X-Signature": naver_signature(ts, "GET", path),
    }
    # 힌트 키워드는 공백 없이, 최대 5개 권장 → 나눠서 호출
    params = {"hintKeywords": ",".join(k.replace(" ", "") for k in hint_keywords), "showDetail": "1"}
    r = requests.get(NAVER_BASE + path, params=params, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json().get("keywordList", [])


def to_int(v):
    """월간검색수는 '< 10' 같은 문자열이 올 수 있음 → 정수화(미만은 5로 근사)"""
    if isinstance(v, int):
        return v
    s = str(v).strip()
    if s.startswith("<"):
        return 5
    try:
        return int(s.replace(",", ""))
    except ValueError:
        return 0


def collect():
    seen = {}
    # 5개씩 끊어서 호출 (API 힌트 상한 대비)
    for i in range(0, len(SEED_KEYWORDS), 5):
        batch = SEED_KEYWORDS[i:i + 5]
        try:
            for row in fetch_keywords(batch):
                kw = row.get("relKeyword")
                if not kw or not is_relevant(kw):
                    continue
                pc = to_int(row.get("monthlyPcQcCnt"))
                mo = to_int(row.get("monthlyMobileQcCnt"))
                comp = row.get("compIdx", "")  # 경쟁정도: 낮음/중간/높음
                total = pc + mo
                # 같은 키워드가 여러 배치에 나오면 최대값 유지
                if kw not in seen or total > seen[kw]["total"]:
                    seen[kw] = {"kw": kw, "pc": pc, "mo": mo, "total": total, "comp": comp}
        except Exception as e:  # noqa: BLE001
            print(f"[warn] batch {batch} 실패: {e}", file=sys.stderr)
        time.sleep(0.3)  # 레이트리밋 여유
    rows = sorted(seen.values(), key=lambda x: x["total"], reverse=True)
    return rows[:80]  # 풀 유지 → 아래에서 '검색량순'과 '경쟁낮은 기회'로 나눠 씀


def week_label(d: datetime.date) -> str:
    nth = (d.day - 1) // 7 + 1
    return f"{d.month}월 {nth}주차 ({d:%m/%d}) · 네이버 (블로그·인스타)"


def _rt(text, bold=False):
    return [{"type": "text", "text": {"content": str(text)[:1900]}, "annotations": {"bold": bold}}]


def _heading(text):
    return {"object": "block", "type": "heading_3", "heading_3": {"rich_text": _rt(text)}}


def _para(text):
    return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": _rt(text)}}


def _bullet(text):
    return {"object": "block", "type": "bulleted_list_item", "bulleted_list_item": {"rich_text": _rt(text)}}


def _todo(text):
    # 체크박스 + 텍스트. 검색어를 콘텐츠에 실제로 쓰면 이 체크박스만 체크해 사용완료 표시.
    return {"object": "block", "type": "to_do", "to_do": {"rich_text": _rt(text), "checked": False}}


def _callout(text, emoji="✍️", color="gray_background"):
    return {"object": "block", "type": "callout",
            "callout": {"icon": {"emoji": emoji}, "color": color, "rich_text": _rt(text)}}


def _toggle(title, children):
    return {"object": "block", "type": "toggle", "toggle": {"rich_text": _rt(title), "children": children[:100]}}


def _table(header, rows):
    def row(cells, bold=False):
        return {"object": "block", "type": "table_row", "table_row": {"cells": [_rt(c, bold) for c in cells]}}
    return {"object": "block", "type": "table",
            "table": {"table_width": len(header), "has_column_header": True, "has_row_header": False,
                      "children": [row(header, True)] + [row(r) for r in rows[:98]]}}


def _line(r):
    return f"{r['kw']}  —  PC {r['pc']:,} / 모바일 {r['mo']:,}  (합 {r['total']:,}, 경쟁 {r['comp']})"


def _row(r):
    return [r["kw"], f"{r['pc']:,}", f"{r['mo']:,}", f"{r['total']:,}", r["comp"] or "-"]


def build_children(rows):
    # 난이도별: 💎 우선 공략(경쟁 낮음·중간) / 🔥 경쟁 치열(높음). 둘 다 검색량순.
    easy = [r for r in rows if r["comp"] != "높음" and r["total"] >= 30][:12]
    hard = [r for r in rows if r["comp"] == "높음"][:8]
    header = ["검색어", "PC", "모바일", "합계 (월간)", "경쟁"]

    ch = [_callout(
        "숫자 읽는 법 — "
        "PC / 모바일: 지난 한 달간 네이버에서 이 검색어를 친 횟수(네이버 검색광고 실데이터). 미용사는 대부분 모바일.  "
        "합계: 수요 크기. 100~1,000이면 롱테일(적지만 정확한 손님), 1,000↑이면 메인 검색어.  "
        "경쟁: 네이버 광고 경쟁도(낮음/중간/높음). 블로그 순위 그 자체는 아니지만 '돈 되는 검색어'일수록 글도 많다는 근사치 — 낮음·중간이 상위 잡기 쉬움.  "
        "💎 = 경쟁 낮음·중간이면서 검색량 있음 / 🔥 = 검색량 크지만 경쟁 높음(장기전).",
        "📐")]
    ch.append(_heading("💎 우선 공략 (경쟁 덜함 · 상위 잡기 유리)"))
    ch.append(_table(header, [_row(r) for r in easy]) if easy else _bullet("해당 없음"))
    ch.append(_heading("🔥 검색량 크지만 경쟁 치열 (장기전 · 참고)"))
    ch.append(_table(header, [_row(r) for r in hard]) if hard else _bullet("해당 없음"))

    main_kw = easy[0]["kw"] if easy else (rows[0]["kw"] if rows else "미용가위")
    subs = [r["kw"] for r in easy[1:3]] or [r["kw"] for r in rows[1:3]]
    ch.append(_heading("🧩 조합 가이드 — 블로그 제목 · 인스타 캡션 만들 때"))
    ch.append(_table(
        ["채널", "자리", "어디서 고르나", "규칙", "이번 주 예시"],
        [
            ["블로그", "제목 앞쪽 + 첫 문단 1~2회", "💎 1개 (메인)", "숫자형 또는 질문형, 낚시 ❌", f"{main_kw} 고르는 기준 3가지"],
            ["블로그", "소제목(H2) 3개", "💎 보조 2개", "소제목마다 1개씩, 억지 반복 ❌", " / ".join(subs) or "-"],
            ["블로그", "태그 5~10개", "메인+보조+카테고리", "보조 수단일 뿐, 태그로 순위 안 오름", f"#{main_kw.replace(' ', '')} …"],
            ["인스타", "캡션 첫 줄(후킹)", "💎 1개", "질문 또는 반전 한 줄", f"{main_kw}, 비싼 게 답일까?"],
            ["인스타", "해시태그 롱테일 1~3개", "💎 표에서 합계 100~1,000짜리", "브랜드 고정 4개 + 카테고리 + 롱테일", "#마모루 #미용가위 + 롱테일"],
            ["인스타 릴스", "첫 1초 화면 텍스트", "💎 1개", "검색어 그대로 짧게", main_kw],
        ]))
    ch.append(_callout(
        "Claude에게 이렇게 요청: \"이번 주 네이버 로그 💎로 블로그 제목 5개 + 인스타 캡션 첫 줄 3개 뽑아줘 (브랜드 톤)\". "
        "🔥는 시리즈로 여러 편 쌓을 때만. 쓴 검색어는 아래 토글의 ☑ 체크로 사용완료 표시.", "✍️"))
    ch.append(_toggle("☑ 사용완료 체크용 (전체 검색어)",
                      [_todo(("💎 " if r in easy else "🔥 " if r in hard else "· ") + _line(r)) for r in (easy + hard)]))
    return ch[:100]


def create_notion_entry(rows):
    today = datetime.date.today()
    top_kw = ", ".join(r["kw"] for r in rows[:7])
    payload = {
        "parent": {"database_id": NOTION_DB_ID},
        "properties": {
            "주차": {"title": [{"text": {"content": week_label(today)}}]},
            "날짜": {"date": {"start": today.isoformat()}},
            "검색어": {"rich_text": [{"text": {"content": top_kw}}]},
            "채널": {"multi_select": [{"name": "블로그"}, {"name": "인스타"}]},
            "반영": {"checkbox": False},
            "메모": {"rich_text": [{"text": {"content": "GitHub Actions 자동 기록 (네이버 검색광고 API · 월간검색량 실데이터)"}}]},
        },
        "children": build_children(rows),
    }
    r = requests.post(
        "https://api.notion.com/v1/pages",
        headers={
            "Authorization": f"Bearer {NOTION_TOKEN}",
            "Notion-Version": "2022-06-28",
            "Content-Type": "application/json",
        },
        json=payload, timeout=30,
    )
    if r.status_code >= 300:
        print(f"[error] Notion 기록 실패 {r.status_code}: {r.text}", file=sys.stderr)
        r.raise_for_status()
    print(f"[ok] Notion에 {len(rows)}개 키워드 기록 완료 — {week_label(today)}")


def main():
    rows = collect()
    if not rows:
        print("[error] 수집된 키워드가 없습니다. 네이버 키/서명/씨앗 키워드를 확인하세요.", file=sys.stderr)
        sys.exit(1)
    create_notion_entry(rows)


if __name__ == "__main__":
    main()
