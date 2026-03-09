#!/usr/bin/env python3
"""
주간 금리 동향 HTML 보고서 생성 스크립트
- 이동찬 연습.xlsx (Sheet2) 에서 데이터 읽기
- HTML 보고서 생성 (차트, 전기 대비 변동폭, 트렌드 코멘트 포함)
"""

import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta
import os
import sys

# ── 설정 ──────────────────────────────────────────────────────────────────────
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '이동찬 연습.xlsx')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')

# ── 섹션 설정 (이름, 아이콘, 헤더행, 데이터행 범위) ─────────────────────────────
SECTION_CONFIGS = [
    ('중앙은행 기준금리',                         '🏦', 1,  (2,  7)),
    ('한국 국채 금리',                            '📊', 8,  (9,  12)),
    ('콜금리·CD·COFIX',                          '💰', 13, (14, 17)),
    ('회사채 3년물',                              '🏢', 18, (19, 21)),
    ('여전채 민평금리 (3년물)',                    '💳', 22, (23, 28)),
    ('기업어음(CP) 금리 (1년물)',                 '📄', 29, (30, 33)),
    ('저축은행 예금금리 (1년, 창구 기본상품 기준)', '🏪', 34, (35, 40)),
]

# ── 색상 팔레트 ────────────────────────────────────────────────────────────────
COLORS = {
    'primary': '#2764EB',
    'secondary': '#4F46E5',
    'accent': '#10B981',
    'warning': '#F59E0B',
    'danger': '#EF4444',
    'bg': '#F5F7FB',
    'card': '#FFFFFF',
    'text': '#1E293B',
    'muted': '#64748B',
    'border': '#E2E8F0',
    'up': '#EF4444',
    'down': '#2764EB',
    'flat': '#64748B',
}

CHART_COLORS = ['#2764EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#14B8A6']

# ── 데이터 로딩 ────────────────────────────────────────────────────────────────
def load_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name='Sheet2', header=None)

    sections = {}

    def parse_section(name, header_row, data_rows_range, label_col=1, date_start_col=2):
        items = {}
        # 날짜 헤더
        dates = []
        for col in range(date_start_col, df.shape[1]):
            val = df.iloc[header_row, col]
            if pd.notna(val):
                if hasattr(val, 'to_pydatetime'):
                    dates.append((col, val.to_pydatetime()))
                else:
                    try:
                        dates.append((col, pd.to_datetime(val)))
                    except:
                        pass

        # 각 지표별 데이터
        for row in range(data_rows_range[0], data_rows_range[1]):
            label = df.iloc[row, label_col]
            if pd.isna(label):
                continue
            label = str(label).replace('\n', ' ').strip()
            series = {}
            for col, dt in dates:
                val = df.iloc[row, col]
                if pd.notna(val):
                    series[dt] = float(val)
            if series:
                items[label] = series

        return {'dates': [d for _, d in dates], 'items': items}

    for name, _icon, header_row, data_rows in SECTION_CONFIGS:
        sections[name] = parse_section(name, header_row, data_rows)

    return sections

# ── 헬퍼 함수 ──────────────────────────────────────────────────────────────────
def get_latest_two(series_dict):
    """가장 최근 두 기간의 데이터 반환 (current, previous, current_date, prev_date)"""
    if not series_dict:
        return None, None, None, None
    sorted_items = sorted(series_dict.items(), key=lambda x: x[0])
    valid = [(d, v) for d, v in sorted_items if pd.notna(v)]
    if len(valid) == 0:
        return None, None, None, None
    curr_date, curr_val = valid[-1]
    if len(valid) >= 2:
        prev_date, prev_val = valid[-2]
    else:
        prev_date, prev_val = None, None
    return curr_val, prev_val, curr_date, prev_date

def change_badge(curr, prev, decimals=3):
    """변동폭 HTML 배지 생성 (bp = diff * 10000)"""
    if curr is None or prev is None:
        return '<span style="color:#94a3b8">-</span>'
    diff = curr - prev
    bp = diff * 10000  # 소수 → bp 변환 (예: 0.0025 → 25bp)
    if abs(bp) < 0.0001:
        arrow = '→'
        color = COLORS['flat']
        sign = ''
    elif bp > 0:
        arrow = '▲'
        color = COLORS['up']
        sign = '+'
    else:
        arrow = '▼'
        color = COLORS['down']
        sign = ''
    return f'<span style="color:{color};font-weight:bold">{arrow} {sign}{bp:.{decimals}f}bp</span>'

SECTIONS_3DP = {'한국 국채 금리', '회사채 3년물', '여전채 민평금리 (3년물)'}

def fmt_date_label(dt):
    """day==1이면 월 단위(YYYY년 MM월), 그 외엔 일자 포함(YYYY년 MM월 DD일)"""
    if dt is None:
        return None
    return dt.strftime('%Y년 %m월') if dt.day == 1 else dt.strftime('%Y년 %m월 %d일')

def fmt_chart_label(dt):
    """차트 x축: day==1이면 YY.MM, 그 외엔 YY.MM.DD"""
    return dt.strftime('%y.%m') if dt.day == 1 else dt.strftime('%y.%m.%d')

def pct_format(val, decimals=2):
    if val is None:
        return '-'
    return f'{val*100:.{decimals}f}%'

def generate_trend_comment(section_name, items):
    """섹션별 간단한 트렌드 코멘트 자동 생성"""
    comments = []
    ups, downs, flats = [], [], []

    for label, series in items.items():
        curr, prev, _, _ = get_latest_two(series)
        if curr is None or prev is None:
            continue
        diff = curr - prev
        if abs(diff) < 0.0001:
            flats.append(label)
        elif diff > 0:
            ups.append((label, diff))
        else:
            downs.append((label, diff))

    if ups:
        top = sorted(ups, key=lambda x: x[1], reverse=True)
        labels_str = ', '.join([f'{l}' for l, _ in top[:2]])
        comments.append(f'<b>{labels_str}</b> 상승')
    if downs:
        top = sorted(downs, key=lambda x: x[1])
        labels_str = ', '.join([f'{l}' for l, _ in top[:2]])
        comments.append(f'<b>{labels_str}</b> 하락')
    if flats:
        labels_str = ', '.join(flats[:2])
        comments.append(f'<b>{labels_str}</b> 보합')

    if not comments:
        return '전기 대비 변동 없음 또는 데이터 미입력 상태입니다.'

    return ' / '.join(comments) + ' (전기 대비)'

# ── 차트 생성 (Chart.js 인터랙티브) ───────────────────────────────────────────
_chart_id = 0

def make_chartjs_html(section_name, items, n_recent=12):
    """Chart.js 인터랙티브 라인 차트 HTML 반환 (호버 툴팁 + y축 숫자 포함)"""
    global _chart_id
    _chart_id += 1
    chart_id = f'chart_{_chart_id}'

    datasets = []
    labels = []

    for i, (label, series) in enumerate(items.items()):
        if not series:
            continue
        sorted_s = sorted(series.items(), key=lambda x: x[0])[-n_recent:]
        if not sorted_s:
            continue

        # x축 레이블은 첫 번째 지표 기준
        if not labels:
            labels = [fmt_chart_label(d) for d, _ in sorted_s]

        color = CHART_COLORS[i % len(CHART_COLORS)]
        # hex → rgba
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)

        data_points = [round(v * 100, 4) for _, v in sorted_s]

        datasets.append({
            'label': label,
            'data': data_points,
            'borderColor': color,
            'backgroundColor': f'rgba({r},{g},{b},0.08)',
            'borderWidth': 2,
            'pointRadius': 4,
            'pointHoverRadius': 7,
            'pointBackgroundColor': color,
            'tension': 0.3,
            'fill': False,
        })

    if not datasets:
        return ''

    chart_data = json.dumps({'labels': labels, 'datasets': datasets}, ensure_ascii=False)

    return f'''
    <div style="margin-top:20px;border-top:1px solid #E2E8F0;padding-top:16px">
      <div style="font-size:12px;color:#94A3B8;margin-bottom:12px">📈 추이 (최근 12개월)</div>
      <div style="position:relative;height:240px">
        <canvas id="{chart_id}"></canvas>
      </div>
    </div>
    <script>
    (function() {{
      var ctx = document.getElementById('{chart_id}').getContext('2d');
      new Chart(ctx, {{
        type: 'line',
        data: {chart_data},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          plugins: {{
            legend: {{
              position: 'bottom',
              labels: {{
                font: {{ size: 12 }},
                color: '#334155',
                usePointStyle: true,
                padding: 16,
              }}
            }},
            tooltip: {{
              backgroundColor: 'rgba(15,23,42,0.85)',
              titleColor: '#94A3B8',
              bodyColor: '#F1F5F9',
              padding: 12,
              cornerRadius: 8,
              callbacks: {{
                label: function(ctx) {{
                  return ' ' + ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(2) + '%';
                }}
              }}
            }}
          }},
          scales: {{
            x: {{
              grid: {{ display: false }},
              ticks: {{ color: '#64748B', font: {{ size: 11 }} }}
            }},
            y: {{
              grid: {{ color: '#E2E8F0', lineWidth: 1 }},
              ticks: {{
                color: '#64748B',
                font: {{ size: 11 }},
                callback: function(val) {{ return val.toFixed(2); }}
              }}
            }}
          }}
        }}
      }});
    }})();
    </script>'''

# ── 주간 이슈 섹션 ────────────────────────────────────────────────────────────
ISSUE_STYLES = {
    '경제': {'icon': '🏛', 'color': '#2764EB', 'bg': '#EFF6FF'},
    '금융': {'icon': '📈', 'color': '#10B981', 'bg': '#ECFDF5'},
    '시장': {'icon': '📊', 'color': '#8B5CF6', 'bg': '#F5F3FF'},
    '정책': {'icon': '📋', 'color': '#F59E0B', 'bg': '#FFFBEB'},
    '해외': {'icon': '🌏', 'color': '#06B6D4', 'bg': '#ECFEFF'},
}

def build_issues_html(issues):
    """issues: [{'category': '경제', 'title': '...', 'summary': '...'}, ...]"""
    if not issues:
        return ''

    cards_html = ''
    for issue in issues:
        cat = issue.get('category', '기타')
        style = ISSUE_STYLES.get(cat, {'icon': '📌', 'color': '#64748B', 'bg': '#F8FAFC'})
        title = issue.get('title', '')
        summary = issue.get('summary', '')

        cards_html += f'''
        <div style="background:{style['bg']};border-radius:12px;padding:16px 20px;
                    border-left:4px solid {style['color']};">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
            <span style="font-size:16px">{style['icon']}</span>
            <span style="font-size:11px;font-weight:700;color:{style['color']};
                         background:white;padding:2px 10px;border-radius:20px;
                         border:1px solid {style['color']};">{cat}</span>
            <span style="font-size:14px;font-weight:700;color:#1E293B">{title}</span>
          </div>
          <p style="font-size:13px;color:#475569;margin:0;line-height:1.6;padding-left:2px">{summary}</p>
        </div>'''

    return f'''
    <div style="background:#fff;border-radius:16px;padding:24px 28px;margin-bottom:24px;
                box-shadow:0 2px 12px rgba(0,0,0,0.06);border:1px solid #E8EEF8">
      <div style="display:flex;align-items:center;margin-bottom:18px;gap:10px">
        <span style="font-size:22px">📰</span>
        <h2 style="margin:0;font-size:17px;font-weight:700;color:#1E293B">주간 주요 이슈</h2>
      </div>
      <div style="display:flex;flex-direction:column;gap:12px">
        {cards_html}
      </div>
    </div>'''

# ── HTML 생성 ──────────────────────────────────────────────────────────────────
def build_section_html(section_name, section_data, icon, rate_decimals=2):
    items = section_data['items']
    if not items:
        return ''

    comment = generate_trend_comment(section_name, items)
    chart_html_block = make_chartjs_html(section_name, items)

    # 모든 섹션 bp 소수점 1자리 (중앙은행·저축은행 금리는 25bp 단위로 움직이므로 의도적)
    bp_decimals = 1

    # 테이블 행
    rows_html = ''
    for label, series in items.items():
        curr, prev, curr_date, prev_date = get_latest_two(series)
        badge = change_badge(curr, prev, decimals=bp_decimals)
        curr_str = pct_format(curr, decimals=rate_decimals)
        prev_str = pct_format(prev, decimals=rate_decimals)
        rows_html += f'''
        <tr>
          <td style="padding:10px 14px;font-weight:600;color:#334155;white-space:nowrap">{label}</td>
          <td style="padding:10px 14px;text-align:right;font-family:monospace;font-size:15px;color:#1E293B;font-weight:700">{curr_str}</td>
          <td style="padding:10px 14px;text-align:right;font-family:monospace;color:#64748B">{prev_str}</td>
          <td style="padding:10px 14px;text-align:right">{badge}</td>
        </tr>'''

    # 날짜 레이블
    sample_series = next(iter(items.values()))
    curr_val, prev_val, curr_date, prev_date = get_latest_two(sample_series)
    curr_label = fmt_date_label(curr_date) or '최신'
    prev_label = fmt_date_label(prev_date) or '전기'

    chart_html = chart_html_block

    return f'''
    <div style="background:#fff;border-radius:16px;padding:24px 28px;margin-bottom:24px;
                box-shadow:0 2px 12px rgba(0,0,0,0.06);border:1px solid #E8EEF8">

      <!-- 섹션 헤더 -->
      <div style="display:flex;align-items:center;margin-bottom:16px;gap:10px">
        <span style="font-size:22px">{icon}</span>
        <h2 style="margin:0;font-size:17px;font-weight:700;color:#1E293B">{section_name}</h2>
      </div>

      <!-- 트렌드 코멘트 -->
      <div style="background:#F0F5FF;border-left:4px solid #2764EB;
                  padding:10px 16px;border-radius:0 8px 8px 0;
                  font-size:13.5px;color:#334155;margin-bottom:18px;line-height:1.6">
        💡 {comment}
      </div>

      <!-- 데이터 테이블 -->
      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:14px">
          <thead>
            <tr style="background:#F1F5F9">
              <th style="padding:10px 14px;text-align:left;color:#64748B;font-weight:600;border-radius:8px 0 0 8px">항목</th>
              <th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600">{curr_label}</th>
              <th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600">{prev_label}</th>
              <th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;border-radius:0 8px 8px 0">변동</th>
            </tr>
          </thead>
          <tbody>
            {rows_html}
          </tbody>
        </table>
      </div>

      {chart_html}
    </div>'''

def generate_html_report(sections, issues=None):
    today = datetime.now()
    week_num = today.isocalendar()[1]
    report_date = today.strftime('%Y년 %m월 %d일')

    sections_html = ''
    for name, icon, _header_row, _data_rows in SECTION_CONFIGS:
        if name in sections:
            decimals = 3 if name in SECTIONS_3DP else 2
            sections_html += build_section_html(name, sections[name], icon, rate_decimals=decimals)

    html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>[웰컴에프앤디] 주간 금리 동향 - {report_date}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Noto Sans KR', 'Malgun Gothic', '맑은 고딕', sans-serif;
    background: #F5F7FB;
    color: #1E293B;
    padding: 0;
  }}
  .wrap {{ max-width: 860px; margin: 0 auto; padding: 32px 20px; }}
  table tr:hover td {{ background: #F8FAFC; }}
</style>
</head>
<body>
<div class="wrap">

  <!-- ── 헤더 ── -->
  <div style="background:linear-gradient(135deg,#2764EB,#4F46E5);
              border-radius:20px;padding:36px 40px;color:#fff;
              margin-bottom:32px;box-shadow:0 8px 32px rgba(39,100,235,0.25)">
    <div style="font-size:13px;background:rgba(255,255,255,0.18);
                display:inline-block;padding:4px 14px;border-radius:20px;
                margin-bottom:14px;letter-spacing:0.5px">
      📅 {report_date} 기준
    </div>
    <h1 style="font-size:28px;font-weight:700;line-height:1.3;margin-bottom:8px">
      주간 금리 동향
    </h1>
    <p style="font-size:14px;opacity:0.85;margin:0">
      웰컴에프앤디(주) 자금팀
    </p>
  </div>

  <!-- ── 요약 배너 ── -->
  <div style="background:#fff;border-radius:14px;padding:18px 24px;
              margin-bottom:28px;border:1px solid #E2E8F0;
              box-shadow:0 2px 8px rgba(0,0,0,0.04)">
    <p style="font-size:13px;color:#64748B;margin-bottom:4px">📌 이 보고서는 주요 금리 지표의 최근 동향을 요약한 자료입니다.</p>
    <p style="font-size:13px;color:#64748B">수치는 <b>전기(전월) 대비</b> 변동폭(bp, basis point)으로 표시됩니다. (1bp = 0.01%)</p>
  </div>

  <!-- ── 주간 이슈 ── -->
  {build_issues_html(issues)}

  <!-- ── 섹션들 ── -->
  {sections_html}

  <!-- ── 푸터 ── -->
  <div style="text-align:center;padding:24px 0;color:#94A3B8;font-size:12px;
              border-top:1px solid #E2E8F0;margin-top:8px">
    <p>웰컴에프앤디(주) 자금팀 | 이 보고서는 내부 참고용으로 작성되었습니다.</p>
    <p style="margin-top:4px">데이터 출처: 금융투자협회 채권정보센터, 은행연합회 소비자포털, 한국은행 경제통계시스템, 세이브로, 저축은행중앙회 소비자포털</p>
  </div>

</div>
</body>
</html>'''

    return html

# ── 메인 실행 ──────────────────────────────────────────────────────────────────
def main():
    print('📊 주간 금리 동향 보고서 생성 시작...')

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(EXCEL_PATH):
        print(f'❌ 엑셀 파일을 찾을 수 없습니다: {EXCEL_PATH}')
        sys.exit(1)

    # 커맨드라인에서 이슈 JSON 받기 (예: python3 script.py '[{"category":"경제",...}]')
    # 인수가 없으면 아래 기본 이슈 사용
    DEFAULT_ISSUES = [
        {
            "category": "해외",
            "title": "미·이란 전쟁 확전 & 유가 급등",
            "summary": "2월 28일 미국·이스라엘의 이란 합동 공습으로 하메네이 최고지도자 사망이 확인되었고, 이란이 걸프 지역 미군 기지에 보복 미사일을 발사하며 전면 충돌로 확산. 이란의 호르무즈 해협 봉쇄 위협에 WTI 원유가 주간 기준 역대 최대폭(+35.6%) 급등하며 글로벌 금융시장 전반에 리스크 오프(Risk-Off) 국면 형성. 금·달러 동반 강세, 주요 증시 급락."
        },
        {
            "category": "시장",
            "title": "국채금리 급등 — 이란 충격·인플레 재평가",
            "summary": "이란發 유가 쇼크로 인플레이션 경로 재평가가 진행되면서 국고채 3년물 금리가 전주 대비 10bp 이상 올라 3.31% 안팎까지 상승, 10년물·30년물도 8~9bp 동반 상승. 미국 10년물 국채 금리 역시 4.05%를 돌파하며 연준의 금리 인하 기대가 후퇴. 전쟁 장기화·유가 80달러 상회 시 금리 상단이 한 단계 더 높아질 가능성 경계."
        },
        {
            "category": "금융",
            "title": "여전채 크레딧 스프레드 확대 압력",
            "summary": "중동 리스크 확대와 국채금리 급등으로 여전채(여신전문금융채) 신용스프레드가 확대 압력을 받고 있음. 2025년 연간 여전채 발행 규모는 95.5조 원으로 전년 대비 약 5% 감소. 크레딧 스프레드(AA- 기준)는 3분기 44bp대에서 연말 53.5bp까지 벌어진 상태로, 이란 사태가 추가 확대 요인으로 부각."
        },
        {
            "category": "경제",
            "title": "저축은행 수신금리 인상 — 머니무브 방어·유동성 관리 목적",
            "summary": "증시 호황에 따른 머니무브(자금 이탈) 가속화로 수신 잔액이 줄어들자, 저축은행들이 선제적 방어 차원에서 예금·파킹통장 금리를 일제히 인상. 대출 정체 속 유동성·예대율 관리 부담이 커진 점도 방어적 금리 인상의 배경. 3~4월 공모주 청약 시즌 단기 대기 자금 유치를 위해 파킹통장 금리와 한도를 공격적으로 확대하는 양상. 향후 증시 변동성 확대 시 대기성 자금 선점 경쟁이 더욱 심화될 전망."
        },
    ]

    issues = DEFAULT_ISSUES
    if len(sys.argv) > 1:
        try:
            issues = json.loads(sys.argv[1])
        except Exception as e:
            print(f'⚠️  이슈 파싱 오류: {e}')

    print('  - 데이터 로딩 중...')
    sections = load_data()

    print('  - HTML 보고서 생성 중...')
    html = generate_html_report(sections, issues=issues)

    today = datetime.now()
    week_num = today.isocalendar()[1]
    filename = f'주간금리동향_{today.strftime("%Y%m%d")}.html'
    output_path = os.path.join(OUTPUT_DIR, filename)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'✅ 보고서 생성 완료: {output_path}')
    return output_path

if __name__ == '__main__':
    main()
