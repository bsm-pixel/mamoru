# -*- coding: utf-8 -*-
"""위젯 43종의 HTML/CSS/JS를 읽어 _design_lab/widgets_preview.html 의
<script id="mm-data"> 블록에 JSON으로 굽는다. 위젯 수정 후 문서 재생성과 함께 실행."""
import re, os, json

HERE = os.path.dirname(os.path.abspath(__file__))
PREVIEW = os.path.normpath(os.path.join(HERE, '..', '_design_lab', 'widgets_preview.html'))

FOLDERS = ['01_before_after','02_countdown','03_compare','04_calculator','05_counter','06_review_carousel','07_care_guide','08_treatment_filter','09_glossary','10_hours_badge','11_recommender','12_360_viewer','13_option_preview','14_cut_hero','15_dealer_gate','16_daily_pick','d01_cinematic_banner','d02_split_promo','d03_notice_bar','d04_quote_banner','d05_dual_choice','n01_size_ruler','n02_versus_table','n03_horizontal_pin','n04_text_mask','n05_stock_gauge','n06_steps','n07_section_header','n08_magnetic_cta','n09_stat_grid','n10_icon_nav','n11_youtube_banner','n12_category_grid','t01_storytelling','t02_repair_timeline','t03_hotspot','t04_tilt_cards','t05_marquee','t06_lightbox_gallery','t07_spotlight','t08_typing','t09_progress','t10_grip_simulator','t11_logo_marquee','n13_event_cards','n14_feature_rows','n15_product_tabs']

def read(p):
    with open(p, encoding='utf-8') as h: return h.read()

def title_of(fid, html):
    m = re.search(r'MAMORU\s*커스텀\s*위젯\s*[—\-–]\s*(.+)', html)
    name = m.group(1).strip() if m else fid
    name = re.sub(r'\s*\(.*$', '', name).strip()  # 괄호 부제 제거
    return fid + '  ·  ' + name

def main():
    data = []
    for fid in FOLDERS:
        d = os.path.join(HERE, fid)
        html = read(os.path.join(d, 'widget.html'))
        css  = read(os.path.join(d, 'widget.css'))
        js   = read(os.path.join(d, 'widget.js'))
        data.append({'id': fid, 'title': title_of(fid, html), 'html': html, 'css': css, 'js': js})

    payload = json.dumps(data, ensure_ascii=False)
    payload = payload.replace('</', '<\\/')  # <script> 조기종료 방지(JSON에선 <\/ == </)

    page = read(PREVIEW)
    new = re.sub(
        r'(<script id="mm-data" type="application/json">)[\s\S]*?(</script>)',
        lambda mm: mm.group(1) + '\n' + payload + '\n' + mm.group(2),
        page, count=1)
    with open(PREVIEW, 'w', encoding='utf-8') as h: h.write(new)
    print('widgets_preview.html: %d widgets baked (%d KB)' % (len(data), len(new)//1024))

if __name__ == '__main__':
    main()
