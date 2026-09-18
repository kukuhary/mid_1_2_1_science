import os
import subprocess
import fitz

def generate_exam_pdf():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    html_path = os.path.join(base_dir, 'exam_15_v1.html')
    pdf_path = os.path.join(base_dir, '2026_중1_2학기중간_과학_실전모의고사_15제_v1.pdf')
    chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
    edge_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

    html_content = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>2026학년도 중1 2학기 중간고사 대비 과학 실전 모의고사 (15제 v1)</title>
<style>
  @page {
    size: A4;
    margin: 12mm 10mm 12mm 10mm;
    @bottom-center {
      content: "- " counter(page) " -";
      font-size: 8.5pt;
      font-family: 'Malgun Gothic', '맑은 고딕', sans-serif;
      color: #333;
    }
  }
  * {
    box-sizing: border-box;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  body {
    font-family: 'Malgun Gothic', '맑은 고딕', 'Batang', '바탕', sans-serif;
    color: #000;
    line-height: 1.72;
    font-size: 9.8pt;
    margin: 0;
    padding: 0;
    background: #fff;
  }

  /* 상단 헤더 박스 */
  .header-box {
    border: 2px solid #000;
    padding: 8px 14px;
    margin-bottom: 12px;
    text-align: center;
    background-color: #fff;
  }
  .header-title {
    font-size: 13.5pt;
    font-weight: bold;
    margin: 0 0 4px 0;
    letter-spacing: -0.5px;
  }
  .header-sub {
    font-size: 8.8pt;
    color: #333;
    margin-bottom: 4px;
  }
  .header-table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 4px;
    font-size: 8.8pt;
  }
  .header-table td {
    border: 1px solid #000;
    padding: 3px 6px;
    text-align: center;
  }
  .header-table .th-bg {
    background-color: #eee;
    font-weight: bold;
  }

  /* 2단 레이아웃 */
  .columns {
    column-count: 2;
    column-gap: 22px;
    column-rule: 1px solid #999;
  }

  .question-item {
    break-inside: avoid;
    margin-bottom: 24px;
    font-size: 9.8pt;
  }
  .question-q {
    font-weight: bold;
    margin-bottom: 5px;
    letter-spacing: -0.3px;
  }
  .q-num {
    font-size: 10.2pt;
    display: inline-block;
    margin-right: 2px;
  }
  .box-view {
    border: 1px solid #444;
    background-color: #fafafa;
    padding: 6px 9px;
    margin: 6px 0;
    font-size: 9.2pt;
    line-height: 1.62;
  }
  .box-title {
    font-weight: bold;
    text-align: center;
    margin-bottom: 3px;
    font-size: 8.8pt;
  }
  .choices {
    margin-top: 6px;
    font-size: 9.5pt;
    line-height: 1.75;
  }
  .choices div {
    margin-bottom: 2px;
  }

  /* 인라인 데이터 표 */
  table.data-table {
    width: 100%;
    border-collapse: collapse;
    margin: 6px 0;
    font-size: 8.8pt;
    text-align: center;
  }
  table.data-table th, table.data-table td {
    border: 1px solid #555;
    padding: 3px 4px;
  }
  table.data-table th {
    background-color: #eaeaea;
  }

  /* SVG 다이어그램 컨테이너 */
  .svg-container {
    display: block;
    margin: 6px auto;
    text-align: center;
    border: 1px solid #ddd;
    background: #fff;
    padding: 4px;
  }
  .caption {
    font-size: 8.3pt;
    color: #444;
    text-align: center;
    margin-top: 2px;
    font-weight: bold;
  }

  /* 페이지 나눔 */
  .page-break {
    page-break-before: always;
  }

  /* 해설 섹션 스타일 */
  .sol-header {
    border: 2px solid #000;
    padding: 7px 12px;
    margin-bottom: 12px;
    text-align: center;
    background-color: #fff;
  }
  .sol-title {
    font-size: 13pt;
    font-weight: bold;
    margin: 0;
  }
  .sol-sub {
    font-size: 8.8pt;
    color: #444;
    margin-top: 3px;
  }

  /* 총괄 요약 표 */
  table.summary-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 14px;
    font-size: 8.4pt;
    text-align: center;
  }
  table.summary-table th, table.summary-table td {
    border: 1px solid #444;
    padding: 3px 3px;
  }
  table.summary-table th {
    background-color: #eee;
    font-weight: bold;
  }

  .sol-item {
    break-inside: avoid;
    margin-bottom: 20px;
    border-bottom: 1px dashed #aaa;
    padding-bottom: 12px;
    font-size: 9.1pt;
    line-height: 1.65;
  }
  .sol-q-title {
    font-size: 9.8pt;
    font-weight: bold;
    margin-bottom: 4px;
    background-color: #eee;
    padding: 3px 6px;
    border-left: 4px solid #000;
  }
  .sol-badge {
    font-size: 8.2pt;
    font-weight: normal;
    background-color: #333;
    color: #fff;
    padding: 1px 5px;
    border-radius: 2px;
    margin-left: 6px;
  }
  .sol-section {
    margin-top: 4px;
  }
  .sol-section-title {
    font-weight: bold;
    color: #000;
    margin-top: 3px;
  }
  .why-box {
    margin-top: 5px;
    background-color: #f8f8f8;
    border-left: 3px solid #666;
    padding: 4px 7px;
    font-size: 8.5pt;
    color: #222;
  }
  .why-title {
    font-weight: bold;
    color: #000;
  }

  ins {
    text-decoration: underline;
    font-weight: bold;
  }
</style>
</head>
<body>

<!-- ================= [PART 1] 실전 문제지 ================= -->
<div class="header-box">
  <div class="header-title">2026학년도 1학년 2학기 중간고사 대비 과학 실전 모의고사 (15제)</div>
  <div class="header-sub">출제 범위: Ⅴ. 힘의 작용 (154~193쪽) &amp; Ⅵ. 기체의 성질 중 기체의 압력 (194~204쪽) | ㈜비상교육 (임태훈 외)</div>
  <table class="header-table">
    <tr>
      <td class="th-bg" style="width: 15%;">과 목</td>
      <td style="width: 20%;">과학 1</td>
      <td class="th-bg" style="width: 15%;">시험 시간</td>
      <td style="width: 15%;">45분</td>
      <td class="th-bg" style="width: 15%;">학번 / 성명</td>
      <td style="width: 20%;">1학년 &nbsp;&nbsp;&nbsp;반 &nbsp;&nbsp;&nbsp;번 &nbsp;&nbsp;성명: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  </table>
</div>

<div class="columns">

  <!-- 문제 1 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">1.</span> 다음 중 과학에서 의미하는 '힘'이 작용하여 물체의 운동 상태나 모양이 변한 경우만을 &lt;보기&gt;에서 있는 대로 고른 것은?</div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. 날아오는 축구공을 가슴으로 트래핑하여 멈추게 하였다.<br>
      ㄴ. 중간고사 과학 시험공부를 열심히 하였더니 머리가 지끈거리고 힘이 들었다.<br>
      ㄷ. 손으로 탄력 있는 고무줄의 양쪽 끝을 잡아당겨 길게 늘였다.<br>
      ㄹ. 친구와 함께 무거운 책상을 들어 올려 복도 끝으로 운반하였다.
    </div>
    <div class="choices">
      <div>① ㄱ, ㄷ, ㄹ</div>
      <div>② ㄱ, ㄴ</div>
      <div>③ ㄴ, ㄷ</div>
      <div>④ ㄴ, ㄹ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ, ㄹ</div>
    </div>
  </div>

  <!-- 문제 2 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">2.</span> 물체에 작용하는 힘의 3요소와 화살표를 이용한 힘의 표현 방법에 대한 설명으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?</div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. 힘의 크기를 나타내는 과학적 표준 단위로는 N(뉴턴)을 사용한다.<br>
      ㄴ. 힘을 화살표로 나타낼 때, 화살표의 길이는 힘의 크기에 비례한다.<br>
      ㄷ. 힘의 크기와 방향이 같더라도 작용점의 위치가 다르면 물체의 회전 여부 등 운동 효과가 달라질 수 있다.<br>
      ㄹ. 힘을 나타내는 화살표의 시작점은 힘이 작용하는 방향을 가리킨다.
    </div>
    <div class="choices">
      <div>① ㄱ, ㄷ</div>
      <div>② ㄱ, ㄴ, ㄷ</div>
      <div>③ ㄱ, ㄴ, ㄹ</div>
      <div>④ ㄴ, ㄷ, ㄹ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ, ㄹ</div>
    </div>
  </div>

  <!-- 문제 3 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">3.</span> 다음은 물체의 운동 상태 (가)~(다)를 나타낸 것이다. 이에 대한 설명으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?</div>
    <div class="box-view">
      (가) 수평한 책상 위에 화분이 가만히 놓여 정지해 있다.<br>
      (나) 공중에서 드론이 일정한 높이를 유지하며 가만히 정지해 있다.<br>
      (다) 마찰이 없는 수평면 위를 컬링 스톤이 일정한 속력으로 곧게 이동하고 있다.
    </div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. (가)에서 화분에 작용하는 중력과 책상이 화분을 떠받치는 힘은 두 힘의 평형 관계이다.<br>
      ㄴ. (나)에서 공중에 멈추어 있는 드론에 작용하는 알짜힘(합력)의 크기는 0 N이다.<br>
      ㄷ. (다)에서 컬링 스톤이 계속 운동하고 있으므로 스톤의 운동 방향으로 일정한 크기의 알짜힘이 작용하고 있다.<br>
      ㄹ. 두 힘이 힘의 평형을 이루기 위해서는 반드시 두 힘의 작용점이 동일한 한 물체에 있어야 한다.
    </div>
    <div class="choices">
      <div>① ㄱ, ㄴ</div>
      <div>② ㄴ, ㄷ</div>
      <div>③ ㄱ, ㄷ, ㄹ</div>
      <div>④ ㄱ, ㄴ, ㄹ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ, ㄹ</div>
    </div>
  </div>

  <!-- 문제 4 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">4.</span> 지구와 달에서 측정한 물체의 질량과 무게에 대한 설명으로 옳은 것은?<br><small style="font-weight:normal;">(단, 달 표면에서의 중력은 지구 표면 중력의 약 1/6이다.)</small></div>
    <div class="choices">
      <div>① 물체의 질량은 측정하는 장소나 중력의 크기에 따라 변한다.</div>
      <div>② 달에서 측정한 어떤 물체의 질량은 지구에서 측정한 질량의 1/6이다.</div>
      <div>③ 지구에서 질량이 60 kg인 우주 비행사의 질량은 달에서도 60 kg으로 변함이 없다.</div>
      <div>④ 지구 표면에서 질량이 1 kg인 물체에 작용하는 중력의 크기(무게)는 약 1 N이다.</div>
      <div>⑤ 달에서 물체의 무게가 줄어드는 까닭은 물체를 구성하는 고유한 물질의 양이 감소하기 때문이다.</div>
    </div>
  </div>

  <!-- 문제 5 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">5.</span> 지구에서 질량이 12 kg인 어떤 물체를 지구와 달에서 각각 서로 다른 저울로 측정하려고 한다. 이에 대한 설명으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?<br><small style="font-weight:normal;">(단, 지구 표면에서 질량 1 kg인 물체의 무게는 9.8 N이고, 달의 중력은 지구의 약 1/6이다.)</small></div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. 지구에서 용수철저울로 측정한 이 물체의 무게는 약 117.6 N이다.<br>
      ㄴ. 달 표면에서 용수철저울로 이 물체의 무게를 측정하면 지구에서 측정한 무게의 1/6인 약 19.6 N이 된다.<br>
      ㄷ. 달 표면에서 양팔저울의 한쪽에 이 물체를 올려놓고 수평을 맞추려면 지구에서 사용했던 것과 동일한 12 kg 분동이 필요하다.
    </div>
    <div class="choices">
      <div>① ㄱ</div>
      <div>② ㄴ</div>
      <div>③ ㄱ, ㄴ</div>
      <div>④ ㄴ, ㄷ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ</div>
    </div>
  </div>

  <!-- 문제 6 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">6.</span> 표는 어떤 용수철에 매단 추의 무게에 따른 용수철의 전체 길이를 측정한 결과이다.</div>
    <table class="data-table">
      <tr>
        <th>추의 무게 (N)</th>
        <td>0 (추 없음)</td>
        <td>2</td>
        <td>4</td>
        <td>6</td>
      </tr>
      <tr>
        <th>용수철의 전체 길이 (cm)</th>
        <td>10</td>
        <td>14</td>
        <td>18</td>
        <td>22</td>
      </tr>
    </table>
    <div class="question-q" style="font-weight:normal; margin-top:4px;">이 용수철에 어떤 물체 A를 매달았더니 용수철의 전체 길이가 26 cm가 되었다. 이에 대한 설명으로 옳은 것은?<br><small>(단, 용수철은 비례 한계를 벗어나지 않는다.)</small></div>
    <div class="choices">
      <div>① 물체 A의 무게는 8 N이다.</div>
      <div>② 용수철의 원래 길이는 14 cm이다.</div>
      <div>③ 용수철의 전체 길이는 매단 추의 무게에 정비례한다.</div>
      <div>④ 물체 A를 매달았을 때 용수철이 늘어난 길이는 26 cm이다.</div>
      <div>⑤ 이 용수철에 무게가 10 N인 물체를 매달면 용수철의 전체 길이는 28 cm가 된다.</div>
    </div>
  </div>

  <!-- 문제 7 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">7.</span> 다음은 수평면 위에서 직육면체 나무도막을 용수철저울로 일정한 속력으로 끌어당기며 마찰력의 크기를 측정한 탐구 실험 (가)~(라)의 조건과 결과이다.</div>
    <div class="box-view">
      (가) 매끄러운 나무판 위, 나무도막 1개 (넓은 면이 닿음) &nbsp;→ 저울 눈금: 2 N<br>
      (나) 매끄러운 나무판 위, 나무도막 1개 (좁은 면이 닿음) &nbsp;→ 저울 눈금: 2 N<br>
      (다) 매끄러운 나무판 위, 나무도막 2개를 위로 포갬 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;→ 저울 눈금: 4 N<br>
      (라) 거친 사포 위, 나무도막 1개 (넓은 면이 닿음) &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;→ 저울 눈금: 5 N
    </div>
    <div class="question-q" style="font-weight:normal; margin-top:4px;">위 실험에 대한 분석으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?</div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. (가)와 (나)를 비교하면 접촉 면적의 넓이는 마찰력의 크기에 영향을 주지 않음을 알 수 있다.<br>
      ㄴ. (가)와 (다)를 비교할 때 일정하게 유지한 통제 변인은 접촉면의 거칠기이다.<br>
      ㄷ. (가)와 (라)를 비교하면 접촉면이 거칠수록 물체에 작용하는 마찰력의 크기가 커짐을 알 수 있다.
    </div>
    <div class="choices">
      <div>① ㄱ</div>
      <div>② ㄴ</div>
      <div>③ ㄱ, ㄴ, ㄷ</div>
      <div>④ ㄱ, ㄷ</div>
      <div>⑤ ㄴ, ㄷ</div>
    </div>
  </div>

  <!-- 문제 8 (SVG 모식도 탑재) -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">8.</span> 그림과 같이 수평한 바닥 위에 무게가 20 N인 상자가 놓여 있다. 이 상자를 용수철저울로 수평 방향 오른쪽으로 일정한 속력으로 끌어당겼더니 용수철저울의 눈금이 6 N을 가리켰다.</div>
    <div class="svg-container">
      <svg width="300" height="90" viewBox="0 0 300 90" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <pattern id="hatch8" width="6" height="6" patternTransform="rotate(45 0 0)" patternUnits="userSpaceOnUse">
            <line x1="0" y1="0" x2="0" y2="6" stroke="#666" stroke-width="1" />
          </pattern>
          <marker id="ar8-left" viewBox="0 0 10 10" refX="2" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
            <path d="M 10 1 L 0 5 L 10 9 z" fill="#000" />
          </marker>
          <marker id="ar8-right" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto">
            <path d="M 0 1 L 10 5 L 0 9 z" fill="#000" />
          </marker>
        </defs>
        <!-- 바닥면 -->
        <line x1="10" y1="70" x2="290" y2="70" stroke="#000" stroke-width="1.8" />
        <rect x="10" y="70" width="280" height="8" fill="url(#hatch8)" />
        <!-- 상자 -->
        <rect x="65" y="28" width="60" height="42" fill="#f0f0f0" stroke="#000" stroke-width="1.8" />
        <text x="95" y="47" font-family="'Malgun Gothic', sans-serif" font-size="9.5" font-weight="bold" text-anchor="middle">상자</text>
        <text x="95" y="61" font-family="'Malgun Gothic', sans-serif" font-size="8.5" text-anchor="middle">(20 N)</text>
        <!-- 마찰력 화살표 -->
        <line x1="65" y1="68" x2="25" y2="68" stroke="#000" stroke-width="1.8" marker-end="url(#ar8-left)" />
        <text x="43" y="60" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">마찰력</text>
        <!-- 저울 연결선 -->
        <line x1="125" y1="49" x2="150" y2="49" stroke="#000" stroke-width="1.5" />
        <rect x="150" y="39" width="65" height="20" fill="#fff" stroke="#000" stroke-width="1.5" />
        <text x="182" y="53" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">저울: 6 N</text>
        <!-- 당기는 힘 -->
        <line x1="215" y1="49" x2="275" y2="49" stroke="#000" stroke-width="1.8" marker-end="url(#ar8-right)" />
        <text x="245" y="41" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">당기는 힘 (6 N)</text>
        <text x="190" y="20" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-style="italic" text-anchor="middle">[일정한 속력으로 운동 →]</text>
      </svg>
      <div class="caption">&lt;상자를 일정한 속력으로 당길 때의 모식도&gt;</div>
    </div>
    <div class="question-q" style="font-weight:normal; margin-top:4px;">이에 대한 설명으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?</div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. 상자가 일정한 속력으로 움직이는 동안 상자에 작용하는 알짜힘의 크기는 0 N이다.<br>
      ㄴ. 상자가 오른쪽으로 미끄러지는 동안 바닥이 상자에 작용하는 마찰력의 크기는 6 N이고, 방향은 왼쪽이다.<br>
      ㄷ. 이 상자 위에 무게 10 N인 벽돌을 올려놓고 일정한 속력으로 끌어당길 때 필요한 힘은 6 N보다 크다.<br>
      ㄹ. 상자를 세워서 바닥과 닿는 면적을 절반으로 줄인 후 일정한 속력으로 끌어당기면 필요한 힘은 3 N으로 줄어든다.
    </div>
    <div class="choices">
      <div>① ㄱ, ㄴ</div>
      <div>② ㄴ, ㄷ</div>
      <div>③ ㄱ, ㄷ, ㄹ</div>
      <div>④ ㄴ, ㄷ, ㄹ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ</div>
    </div>
  </div>

  <!-- 문제 9 (정밀 SVG 모식도 탑재) -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">9.</span> 그림은 공기 중에서 무게가 10 N인 금속 원통을 용수철저울에 매달아 비커에 담긴 물속에 서서히 넣으면서 저울의 눈금을 측정한 과정을 나타낸 모식도이다.</div>
    <div class="svg-container">
      <svg width="310" height="155" viewBox="0 0 310 155" xmlns="http://www.w3.org/2000/svg">
        <!-- (가) 공기 중 -->
        <g transform="translate(8, 0)">
          <text x="32" y="14" font-family="'Malgun Gothic', sans-serif" font-size="8.8" font-weight="bold" text-anchor="middle">(가) 공기 중</text>
          <rect x="17" y="22" width="30" height="19" fill="#fff" stroke="#000" stroke-width="1.3" />
          <text x="32" y="35" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">10 N</text>
          <line x1="32" y1="41" x2="32" y2="58" stroke="#000" stroke-width="1.2" />
          <rect x="23" y="58" width="18" height="28" fill="#ccc" stroke="#000" stroke-width="1.3" />
          <text x="32" y="75" font-family="'Malgun Gothic', sans-serif" font-size="7.8" text-anchor="middle">원통</text>
        </g>
        <!-- (나) 부피 1/2 잠김 -->
        <g transform="translate(84, 0)">
          <text x="32" y="14" font-family="'Malgun Gothic', sans-serif" font-size="8.8" font-weight="bold" text-anchor="middle">(나) 1/2 잠김</text>
          <rect x="9" y="70" width="46" height="65" fill="#fff" stroke="#000" stroke-width="1.2" />
          <rect x="10" y="82" width="44" height="52" fill="#f2f2f2" />
          <line x1="10" y1="82" x2="54" y2="82" stroke="#555" stroke-dasharray="2,2" stroke-width="1" />
          <rect x="17" y="22" width="30" height="19" fill="#fff" stroke="#000" stroke-width="1.3" />
          <text x="32" y="35" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">8 N</text>
          <line x1="32" y1="41" x2="32" y2="68" stroke="#000" stroke-width="1.2" />
          <rect x="23" y="68" width="18" height="28" fill="#ccc" stroke="#000" stroke-width="1.3" />
          <text x="32" y="146" font-family="'Malgun Gothic', sans-serif" font-size="7.8" text-anchor="middle">수면</text>
        </g>
        <!-- (다) 완전히 잠김 -->
        <g transform="translate(160, 0)">
          <text x="32" y="14" font-family="'Malgun Gothic', sans-serif" font-size="8.8" font-weight="bold" text-anchor="middle">(다) 완전 잠김</text>
          <rect x="9" y="70" width="46" height="65" fill="#fff" stroke="#000" stroke-width="1.2" />
          <rect x="10" y="80" width="44" height="54" fill="#f2f2f2" />
          <line x1="10" y1="80" x2="54" y2="80" stroke="#555" stroke-dasharray="2,2" stroke-width="1" />
          <rect x="17" y="22" width="30" height="19" fill="#fff" stroke="#000" stroke-width="1.3" />
          <text x="32" y="35" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">A N</text>
          <line x1="32" y1="41" x2="32" y2="88" stroke="#000" stroke-width="1.2" />
          <rect x="23" y="88" width="18" height="28" fill="#ccc" stroke="#000" stroke-width="1.3" />
        </g>
        <!-- (라) 완전히 잠긴 채 더 깊이 -->
        <g transform="translate(236, 0)">
          <text x="32" y="14" font-family="'Malgun Gothic', sans-serif" font-size="8.8" font-weight="bold" text-anchor="middle">(라) 더 깊은 곳</text>
          <rect x="9" y="70" width="46" height="75" fill="#fff" stroke="#000" stroke-width="1.2" />
          <rect x="10" y="80" width="44" height="64" fill="#f2f2f2" />
          <line x1="10" y1="80" x2="54" y2="80" stroke="#555" stroke-dasharray="2,2" stroke-width="1" />
          <rect x="17" y="22" width="30" height="19" fill="#fff" stroke="#000" stroke-width="1.3" />
          <text x="32" y="35" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">B N</text>
          <line x1="32" y1="41" x2="32" y2="105" stroke="#000" stroke-width="1.2" />
          <rect x="23" y="105" width="18" height="28" fill="#ccc" stroke="#000" stroke-width="1.3" />
        </g>
      </svg>
      <div class="caption">&lt;잠긴 부피 및 수심에 따른 저울 눈금 측정 장치&gt;</div>
    </div>
    <div class="question-q" style="font-weight:normal; margin-top:4px;">이에 대한 분석으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?<br><small>(단, 원통은 비커 바닥이나 벽면에 닿지 않는다.)</small></div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. (나)에서 물이 원통에 작용하는 부력의 크기는 2 N이다.<br>
      ㄴ. (다)에서 용수철저울의 눈금 A는 6 N이다.<br>
      ㄷ. (라)에서 용수철저울의 눈금 B는 A보다 작다.<br>
      ㄹ. 원통이 물에 완전히 잠긴 후에는 수심이 깊어지더라도 원통이 밀어낸 물의 부피가 같으므로 작용하는 부력의 크기는 변하지 않는다.
    </div>
    <div class="choices">
      <div>① ㄱ, ㄷ</div>
      <div>② ㄱ, ㄴ, ㄹ</div>
      <div>③ ㄴ, ㄷ</div>
      <div>④ ㄴ, ㄷ, ㄹ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ, ㄹ</div>
    </div>
  </div>

  <!-- 문제 10 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">10.</span> 다음 (가)~(라)의 운동 사례를 운동 상태(속력과 운동 방향)의 변화 유형에 따라 올바르게 짝지은 것은?</div>
    <div class="box-view">
      (가) 비행기에서 뛰어내린 스카이다이버가 낙하산을 펴기 전 아래로 점점 빨라지며 떨어지는 운동<br>
      (나) 일정한 속력으로 원형 레일을 회전하는 대관람차 곤돌라의 운동<br>
      (다) 비스듬히 차올린 축구공이 포물선을 그리며 골대를 향해 날아가는 운동<br>
      (라) 공항에서 여행 가방을 싣고 일정한 빠르기로 곧게 이동하는 수평 컨베이어 벨트의 운동
    </div>
    <table class="data-table">
      <tr>
        <th></th>
        <th>속력만 변함</th>
        <th>방향만 변함</th>
        <th>속력·방향 모두 변함</th>
        <th>모두 일정함</th>
      </tr>
      <tr>
        <td>①</td>
        <td>(가)</td>
        <td>(나)</td>
        <td>(다)</td>
        <td>(라)</td>
      </tr>
      <tr>
        <td>②</td>
        <td>(가)</td>
        <td>(다)</td>
        <td>(나)</td>
        <td>(라)</td>
      </tr>
      <tr>
        <td>③</td>
        <td>(나)</td>
        <td>(가)</td>
        <td>(라)</td>
        <td>(다)</td>
      </tr>
      <tr>
        <td>④</td>
        <td>(다)</td>
        <td>(나)</td>
        <td>(가)</td>
        <td>(라)</td>
      </tr>
      <tr>
        <td>⑤</td>
        <td>(라)</td>
        <td>(나)</td>
        <td>(다)</td>
        <td>(가)</td>
      </tr>
    </table>
  </div>

  <!-- 문제 11 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">11.</span> 마찰과 공기 저항이 없는 매끄러운 수평면 위에서 어떤 물체가 일정한 속력으로 오른쪽으로 곧게 이동하고 있다. 이 물체의 운동과 작용하는 힘에 대한 설명으로 옳은 것은?</div>
    <div class="choices">
      <div>① 물체의 속력을 일정하게 유지하기 위해서는 오른쪽으로 일정한 크기의 힘을 계속 가해야 한다.</div>
      <div>② 물체가 오른쪽으로 이동하고 있으므로 물체에 작용하는 알짜힘의 방향은 오른쪽이다.</div>
      <div>③ 물체의 운동 방향과 정반대인 왼쪽으로 알짜힘이 작용하면 물체는 즉시 운동 방향을 바꾸어 왼쪽으로 이동한다.</div>
      <div>④ 물체의 운동 방향과 수직인 방향으로 일정한 힘이 계속 작용하면 물체의 속력이 점점 빨라진다.</div>
      <div>⑤ 물체에 작용하는 알짜힘이 0 N일 때 물체는 정지해 있거나 이처럼 일정한 속력으로 곧게 운동한다.</div>
    </div>
  </div>

  <!-- 문제 12 (정밀 SVG 모식도 탑재) -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">12.</span> 그림은 공기 저항이 없는 진공 상태에서 같은 높이에 가만히 정지해 있던 질량 1 kg인 쇠구슬 A와 질량 2 kg인 쇠구슬 B를 동시에 가만히 놓았을 때, 일정한 시간 간격(0.1초 간격)으로 두 쇠구슬의 위치를 연속 촬영한 다중 섬광 사진을 나타낸 모식도이다.</div>
    <div class="svg-container">
      <svg width="300" height="150" viewBox="0 0 300 150" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <marker id="fall-ar12" viewBox="0 0 10 10" refX="5" refY="8" markerWidth="6" markerHeight="6" orient="auto">
            <path d="M 2 0 L 5 9 L 8 0 z" fill="#000" />
          </marker>
        </defs>
        <!-- 배경 헤더 -->
        <rect x="5" y="4" width="290" height="22" fill="#eee" stroke="#666" stroke-width="1" />
        <text x="40" y="19" font-family="'Malgun Gothic', sans-serif" font-size="9" font-weight="bold" text-anchor="middle">경과 시간</text>
        <text x="125" y="19" font-family="'Malgun Gothic', sans-serif" font-size="9" font-weight="bold" text-anchor="middle">쇠구슬 A (1 kg)</text>
        <text x="215" y="19" font-family="'Malgun Gothic', sans-serif" font-size="9" font-weight="bold" text-anchor="middle">쇠구슬 B (2 kg)</text>

        <!-- 세로 구분선 -->
        <line x1="75" y1="4" x2="75" y2="145" stroke="#ccc" stroke-dasharray="2,2" stroke-width="1" />
        <line x1="170" y1="4" x2="170" y2="145" stroke="#ccc" stroke-dasharray="2,2" stroke-width="1" />
        <line x1="260" y1="4" x2="260" y2="145" stroke="#ccc" stroke-dasharray="2,2" stroke-width="1" />

        <!-- 0.0초 -->
        <text x="40" y="42" font-family="'Malgun Gothic', sans-serif" font-size="8.8" text-anchor="middle">0.0초</text>
        <circle cx="125" cy="38" r="4.5" fill="#000" />
        <circle cx="215" cy="38" r="7.5" fill="#888" stroke="#000" stroke-width="1.2" />

        <!-- 0.1초 -->
        <text x="40" y="58" font-family="'Malgun Gothic', sans-serif" font-size="8.8" text-anchor="middle">0.1초</text>
        <circle cx="125" cy="54" r="4.5" fill="#000" />
        <circle cx="215" cy="54" r="7.5" fill="#888" stroke="#000" stroke-width="1.2" />

        <!-- 0.2초 -->
        <text x="40" y="90" font-family="'Malgun Gothic', sans-serif" font-size="8.8" text-anchor="middle">0.2초</text>
        <circle cx="125" cy="86" r="4.5" fill="#000" />
        <circle cx="215" cy="86" r="7.5" fill="#888" stroke="#000" stroke-width="1.2" />

        <!-- 0.3초 -->
        <text x="40" y="136" font-family="'Malgun Gothic', sans-serif" font-size="8.8" text-anchor="middle">0.3초</text>
        <circle cx="125" cy="132" r="4.5" fill="#000" />
        <circle cx="215" cy="132" r="7.5" fill="#888" stroke="#000" stroke-width="1.2" />

        <!-- 낙하 화살표 -->
        <line x1="275" y1="38" x2="275" y2="135" stroke="#000" stroke-width="1.5" marker-end="url(#fall-ar12)" />
        <text x="288" y="90" font-family="'Malgun Gothic', sans-serif" font-size="8" writing-mode="vertical-rl" text-anchor="middle">낙하 방향</text>
      </svg>
      <div class="caption">&lt;진공 상태에서 자유 낙하하는 쇠구슬의 다중 섬광 모식도&gt;</div>
    </div>
    <div class="question-q" style="font-weight:normal; margin-top:4px;">낙하하는 두 쇠구슬의 운동에 대한 설명으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?</div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. 낙하하는 동안 두 쇠구슬 A와 B에 작용하는 중력의 크기는 서로 같다.<br>
      ㄴ. 같은 시간 간격 동안 이동한 거리가 점점 늘어나는 것으로 보아 두 쇠구슬의 속력은 일정하게 빨라진다.<br>
      ㄷ. 두 쇠구슬의 질량은 서로 다르지만 공기 저항이 없으므로 바닥에 동시에 도달한다.<br>
      ㄹ. 두 쇠구슬이 낙하하는 동안 운동 방향과 같은 방향(연직 아래쪽)으로 중력이 계속 작용한다.
    </div>
    <div class="choices">
      <div>① ㄱ, ㄴ</div>
      <div>② ㄴ, ㄷ, ㄹ</div>
      <div>③ ㄱ, ㄷ, ㄹ</div>
      <div>④ ㄴ, ㄷ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ, ㄹ</div>
    </div>
  </div>

  <!-- 문제 13 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">13.</span> 일정한 면적에 작용하는 힘을 '압력'이라고 한다. 압력에 대한 설명 및 일상생활 속 현상으로 옳은 것은?</div>
    <div class="choices">
      <div>① 힘이 작용하는 면적이 넓어질수록 바닥이 받는 압력은 커진다.</div>
      <div>② 작용하는 힘의 크기가 작아질수록 물체가 받는 압력은 커진다.</div>
      <div>③ 눈밭에서 스노슈(설피)를 신으면 바닥에 닿는 면적이 넓어져 눈에 덜 빠지게 된다.</div>
      <div>④ 과도를 사용할 때 칼날을 무디게 갈수록 면적이 좁아져 과일이 더 잘 깎인다.</div>
      <div>⑤ 같은 무게의 직육면체 벽돌이라도 좁은 면이 바닥에 닿을 때 바닥이 받는 압력이 더 작다.</div>
    </div>
  </div>

  <!-- 문제 14 (SVG 일러스트 탑재) -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">14.</span> 다음은 페트병과 쇠구슬을 이용하여 기체의 압력이 나타나는 까닭을 탐구하기 위한 실험 과정이다.</div>
    <div class="svg-container">
      <svg width="290" height="85" viewBox="0 0 290 85" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <marker id="ar14-l" viewBox="0 0 10 10" refX="2" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse">
            <path d="M 10 1 L 0 5 L 10 9 z" fill="#000" />
          </marker>
          <marker id="ar14-r" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto">
            <path d="M 0 1 L 10 5 L 0 9 z" fill="#000" />
          </marker>
        </defs>
        <!-- 페트병 (가) -->
        <g transform="translate(10, 2)">
          <rect x="12" y="10" width="100" height="48" rx="8" fill="#fff" stroke="#000" stroke-width="1.4" />
          <rect x="2" y="24" width="10" height="20" fill="#555" stroke="#000" stroke-width="1" />
          <text x="62" y="26" font-family="'Malgun Gothic', sans-serif" font-size="9" font-weight="bold" text-anchor="middle">페트병 (가)</text>
          <text x="62" y="40" font-family="'Malgun Gothic', sans-serif" font-size="8.5" text-anchor="middle">쇠구슬 15개</text>
          <text x="62" y="52" font-family="'Malgun Gothic', sans-serif" font-size="7.5" fill="#555" text-anchor="middle">(충돌 적음)</text>
          <line x1="40" y1="68" x2="85" y2="68" stroke="#000" stroke-width="1.2" marker-start="url(#ar14-l)" marker-end="url(#ar14-r)" />
          <text x="62" y="80" font-family="'Malgun Gothic', sans-serif" font-size="7.5" text-anchor="middle">좌우로 흔듦</text>
        </g>
        <!-- 페트병 (나) -->
        <g transform="translate(155, 2)">
          <rect x="12" y="10" width="100" height="48" rx="8" fill="#f4f4f4" stroke="#000" stroke-width="1.4" />
          <rect x="2" y="24" width="10" height="20" fill="#555" stroke="#000" stroke-width="1" />
          <text x="62" y="26" font-family="'Malgun Gothic', sans-serif" font-size="9" font-weight="bold" text-anchor="middle">페트병 (나)</text>
          <text x="62" y="40" font-family="'Malgun Gothic', sans-serif" font-size="8.5" font-weight="bold" text-anchor="middle">쇠구슬 30개</text>
          <text x="62" y="52" font-family="'Malgun Gothic', sans-serif" font-size="7.5" fill="#555" text-anchor="middle">(충돌 많음)</text>
          <line x1="40" y1="68" x2="85" y2="68" stroke="#000" stroke-width="1.2" marker-start="url(#ar14-l)" marker-end="url(#ar14-r)" />
          <text x="62" y="80" font-family="'Malgun Gothic', sans-serif" font-size="7.5" text-anchor="middle">좌우로 흔듦</text>
        </g>
      </svg>
      <div class="caption">&lt;페트병과 쇠구슬을 이용한 기체 압력 입자 모형 실험&gt;</div>
    </div>
    <div class="question-q" style="font-weight:normal; margin-top:4px;">위 탐구 실험과 기체의 압력에 대한 설명으로 <ins>옳지 않은</ins> 것은?</div>
    <div class="choices">
      <div>① 페트병 속 쇠구슬은 기체 물질을 구성하는 입자를 나타낸 모형이다.</div>
      <div>② 쇠구슬이 페트병 벽에 부딪히는 것은 기체 입자가 용기 벽에 충돌하는 현상에 해당한다.</div>
      <div>③ 손바닥에 부딪히며 느껴지는 힘은 기체 입자가 충돌하여 나타나는 기체의 압력을 의미한다.</div>
      <div>④ 페트병 (가)가 페트병 (나)보다 손바닥에 부딪히는 쇠구슬의 수가 더 많아 손바닥에 더 큰 힘이 느껴진다.</div>
      <div>⑤ 기체 입자는 모든 방향으로 끊임없이 운동하므로 기체의 압력은 모든 방향으로 작용한다.</div>
    </div>
  </div>

  <!-- 문제 15 -->
  <div class="question-item">
    <div class="question-q"><span class="q-num">15.</span> 다음은 기체의 압력(기압)과 관련된 일상생활 속 현상 (가)~(라)를 나타낸 것이다.</div>
    <div class="box-view">
      (가) 뾰족한 누름못 1개 위에 풍선을 누르면 터지지만, 누름못 30개를 촘촘히 꽂은 판 위에 풍선을 올려놓고 누르면 풍선이 터지지 않는다.<br>
      (나) 매끄러운 유리창에 흡착 고무(빨판)를 대고 꾹 누르면 내부의 공기가 빠져나가 벽에 단단히 달라붙는다.<br>
      (다) 과자 봉지를 가지고 높은 산에 올라가면 과자 봉지가 사방으로 빵빵하게 부풀어 오른다.<br>
      (라) 공기를 가득 넣은 축구공은 특정 방향으로 찌그러지지 않고 사방으로 팽팽하게 둥근 모양을 유지한다.
    </div>
    <div class="question-q" style="font-weight:normal; margin-top:4px;">현상 (가)~(라)에 대한 과학적 분석으로 옳은 것만을 &lt;보기&gt;에서 있는 대로 고른 것은?</div>
    <div class="box-view">
      <div class="box-title">&lt;보 기&gt;</div>
      ㄱ. (가)에서 누름못의 개수가 많아질수록 풍선과 접촉하는 총면적이 넓어져 풍선이 받는 압력이 작아지기 때문이다.<br>
      ㄴ. (나)에서 흡착 고무 내부를 누르면 안쪽 공기가 빠져나가 내부 압력이 낮아지고, 외부 기압이 흡착 고무를 벽 쪽으로 밀어붙이기 때문이다.<br>
      ㄷ. (다)에서 높은 산으로 올라갈수록 공기의 양이 줄어들어 외부 기압이 낮아지기 때문이다.<br>
      ㄹ. (라)에서 축구공이 둥근 모양을 유지하는 까닭은 내부 공기 입자가 오직 바깥쪽 한 방향으로만 운동하기 때문이다.
    </div>
    <div class="choices">
      <div>① ㄱ, ㄷ</div>
      <div>② ㄴ, ㄹ</div>
      <div>③ ㄱ, ㄴ, ㄹ</div>
      <div>④ ㄱ, ㄴ, ㄷ</div>
      <div>⑤ ㄱ, ㄴ, ㄷ, ㄹ</div>
    </div>
  </div>

</div><!-- end columns -->

<!-- ================= [PART 2 & 3] 정답 및 해설지 ================= -->
<div class="page-break"></div>

<div class="sol-header">
  <div class="sol-title">2026학년도 과학 중간고사 대비 실전 모의고사 [정답 및 상세 해설집]</div>
  <div class="sol-sub">난이도 총괄표 및 전 문항 [정답 분석] · [오답 피하기] · [문제를 낸 이유] 완벽 수록</div>
</div>

<!-- 빠른 정답표 및 난이도 총괄표 -->
<table class="summary-table">
  <tr>
    <th style="width: 7%;">문항</th>
    <th style="width: 7%;">정답</th>
    <th style="width: 9%;">난이도</th>
    <th style="width: 22%;">단원 영역</th>
    <th>핵심 평가 요소 및 출제 의도</th>
  </tr>
  <tr>
    <td><b>1</b></td>
    <td><b>①</b></td>
    <td>[하]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>과학에서의 힘(모양·운동 상태 변화) vs 일상어 힘 구별</td>
  </tr>
  <tr>
    <td><b>2</b></td>
    <td><b>②</b></td>
    <td>[중]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>힘의 3요소(작용점, 방향, 크기)와 화살표 표현법</td>
  </tr>
  <tr>
    <td><b>3</b></td>
    <td><b>④</b></td>
    <td>[상]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>두 힘의 평형 3대 조건(한 물체 작용) 및 등속 운동의 알짜힘</td>
  </tr>
  <tr>
    <td><b>4</b></td>
    <td><b>③</b></td>
    <td>[하]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>중력의 정의 및 장소에 따른 질량(불변)과 무게(1/6) 구별</td>
  </tr>
  <tr>
    <td><b>5</b></td>
    <td><b>⑤</b></td>
    <td>[중]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>측정 저울(용수철저울 vs 양팔저울)에 따른 달에서의 측정값 분석</td>
  </tr>
  <tr>
    <td><b>6</b></td>
    <td><b>①</b></td>
    <td>[상]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>용수철의 '늘어난 길이' 비례식 계산 및 미지수 물체 무게 도출</td>
  </tr>
  <tr>
    <td><b>7</b></td>
    <td><b>③</b></td>
    <td>[중]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>마찰력 탐구 실험에서의 변인 통제(무게, 거칠기, 접촉 면적)</td>
  </tr>
  <tr>
    <td><b>8</b></td>
    <td><b>⑤</b></td>
    <td>[상]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>등속 이동 시 알짜힘=0과 마찰력의 크기 및 접촉 면적 무관성</td>
  </tr>
  <tr>
    <td><b>9</b></td>
    <td><b>②</b></td>
    <td>[상]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>부력의 방향, 잠긴 부피 비례성 및 수심 깊이 무관성 검증</td>
  </tr>
  <tr>
    <td><b>10</b></td>
    <td><b>①</b></td>
    <td>[중]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>힘의 방향에 따른 운동 상태 변화 4가지 유형 분류</td>
  </tr>
  <tr>
    <td><b>11</b></td>
    <td><b>⑤</b></td>
    <td>[상]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>등속 직선 운동 시 알짜힘=0 오개념 저격 및 운동 상태 관계</td>
  </tr>
  <tr>
    <td><b>12</b></td>
    <td><b>②</b></td>
    <td>[상]</td>
    <td>Ⅴ. 힘의 작용</td>
    <td>자유 낙하 운동 다중 섬광 사진 해석 및 질량별 중력 크기 구별</td>
  </tr>
  <tr>
    <td><b>13</b></td>
    <td><b>③</b></td>
    <td>[하]</td>
    <td>Ⅵ. 기체의 압력</td>
    <td>일정한 면적에 작용하는 힘(압력)의 정의와 일상생활 속 예</td>
  </tr>
  <tr>
    <td><b>14</b></td>
    <td><b>④</b></td>
    <td>[중]</td>
    <td>Ⅵ. 기체의 압력</td>
    <td>페트병 쇠구슬 모형을 통한 기체 입자 충돌과 압력 발생 원리</td>
  </tr>
  <tr>
    <td><b>15</b></td>
    <td><b>④</b></td>
    <td>[상]</td>
    <td>Ⅵ. 기체의 압력</td>
    <td>기압의 모든 방향 작용성 및 높이에 따른 기압 변화 현상 분석</td>
  </tr>
</table>

<div class="columns">

  <!-- 해설 1 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 1 해설] 정답 ① <span class="sol-badge">난이도: 하</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: 축구공의 속력이 감소하여 멈추었으므로 물체의 <b>운동 상태(속력)가 변한 경우</b>입니다.<br>
      • <b>ㄷ (참)</b>: 고무줄을 잡아당겨 길이가 늘어났으므로 물체의 <b>모양이 변한 경우</b>입니다.<br>
      • <b>ㄹ (참)</b>: 정지해 있던 책상을 들어 올려 다른 곳으로 옮겼으므로 물체의 <b>운동 상태(위치 및 속력)가 변한 경우</b>입니다.<br>
      따라서 과학에서의 힘이 작용한 사례는 ㄱ, ㄷ, ㄹ입니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>ㄴ (거짓)</b>: 시험공부를 열심히 하여 머리가 아프고 힘이 들었다는 것은 피로감이나 정신적 노력을 나타내는 <b>일상생활 속 관용적 표현</b>일 뿐, 물체의 모양이나 운동 상태를 변화시키는 과학에서의 힘이 아닙니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-01] 힘의 작용으로 나타나는 현상을 관찰하고 과학에서의 힘을 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 일상생활의 주관적 피로감·능력으로서의 '힘'과 과학에서 정의하는 객관적 물리량인 '힘(모양 및 운동 상태 변화의 원인)'을 명확히 구분할 수 있는지 확인하고자 출제하였습니다.
    </div>
  </div>

  <!-- 해설 2 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 2 해설] 정답 ② <span class="sol-badge">난이도: 중</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: 과학에서 힘의 크기를 나타내는 표준 단위는 N(뉴턴)입니다. (kg은 질량의 단위)<br>
      • <b>ㄴ (참)</b>: 힘을 화살표로 나타낼 때 화살표의 길이는 힘의 크기에 비례하여 그립니다.<br>
      • <b>ㄷ (참)</b>: 책상 위의 긴 막대를 밀 때 한가운데를 밀면 곧게 앞으로 직진하지만, 끝 모서리를 밀면 물체가 회전합니다. 이처럼 힘의 크기와 방향이 같더라도 <b>작용점의 위치가 다르면 물체에 미치는 운동 효과가 달라집니다.</b>
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>ㄹ (거짓)</b>: 화살표의 <b>시작점은 힘이 작용하는 지점(작용점)</b>이며, 힘이 작용하는 방향을 나타내는 것은 화살표의 <b>화살촉(머리)</b>입니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-01] 힘의 3요소(작용점, 방향, 크기)를 이해하고 화살표로 표현할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 작용점, 방향, 크기의 대응 관계를 점검하고, 작용점의 위치에 따라 회전 여부 등 운동 효과가 결정된다는 교과서 핵심 개념을 체득했는지 평가하기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 3 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 3 해설] 정답 ④ <span class="sol-badge">난이도: 상</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: 정지한 화분은 아래쪽 중력과 위쪽 책상의 수직항력이 크기가 같고 방향이 반대이며 동일 작용선상에 있어 <b>두 힘의 평형</b>을 이룹니다.<br>
      • <b>ㄴ (참)</b>: 드론이 공중에서 일정한 높이를 유지하며 가만히 정지해 있으므로 드론에 작용하는 <b>알짜힘(합력)은 0 N</b>입니다.<br>
      • <b>ㄹ (참)</b>: 두 힘의 평형은 반드시 <b>동일한 '한 물체'에 동시에 작용</b>해야 성립합니다. (작용점이 서로 다른 두 물체에 존재하는 힘은 작용·반작용으로 평형 불가)
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>ㄷ (거짓 - ★최빈출 오개념)</b>: 마찰이 없는 수평면에서 일정한 속력으로 곧게 이동하는 등속 직선 운동을 할 때, 물체에 작용하는 <b>알짜힘은 0 N</b>입니다. 물체가 움직인다고 해서 운동 방향으로 힘이 계속 작용하고 있다고 착각해서는 안 됩니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-02] 두 힘의 합력을 구하고 힘의 평형 조건을 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 중1 학생들이 가장 많이 틀리는 "운동하는 물체에는 반드시 힘이 작용한다"는 오개념을 타파하고, 등속 직선 운동과 정지 상태 모두 알짜힘이 0 N임을 확립하기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 4 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 4 해설] 정답 ③ <span class="sol-badge">난이도: 하</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>③ (참)</b>: <b>질량</b>은 장소나 환경에 관계없이 변하지 않는 <b>물질의 고유한 양</b>입니다. 따라서 지구에서 질량이 60 kg인 우주 비행사는 달에 가더라도 질량이 60 kg으로 변함이 없습니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>① (거짓)</b>: 장소에 따라 변하는 것은 중력의 크기인 '무게'이며, 물질의 고유한 양인 '질량'은 어디서나 일정합니다.<br>
      • <b>② (거짓)</b>: 달에서 1/6로 줄어드는 것은 '무게'이며 질량은 불변입니다.<br>
      • <b>④ (거짓)</b>: 지구 표면에서 질량 1 kg인 물체의 무게는 약 9.8 N입니다.<br>
      • <b>⑤ (거짓)</b>: 달에서 무게가 줄어드는 것은 달이 당기는 중력이 지구의 약 1/6이기 때문이지, 물질의 양이 줄어들어서가 아닙니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-03] 중력의 개념을 이해하고, 무게와 질량의 차이를 비교할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 일상생활의 어휘 혼용을 바로잡고, 불변량인 질량과 중력에 따른 힘인 무게를 명확히 구분할 수 있는지 점검하기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 5 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 5 해설] 정답 ⑤ <span class="sol-badge">난이도: 중</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: 지구에서 질량 12 kg인 물체의 무게는 12 × 9.8 = 117.6 N입니다.<br>
      • <b>ㄴ (참)</b>: 용수철저울은 무게(중력의 크기)를 측정하므로, 중력이 1/6인 달 표면에서는 117.6 ÷ 6 = 약 19.6 N을 가리킵니다.<br>
      • <b>ㄷ (참)</b>: 양팔저울은 양쪽에 작용하는 중력을 비교하는 도구입니다. 달에 가면 물체와 분동에 작용하는 중력이 똑같이 1/6로 감소하여 서로 상쇄되므로, 지구에서와 동일하게 12 kg 분동이 필요합니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • ㄱ, ㄴ, ㄷ 모두 과학적 사실에 부합하므로 정답은 ⑤입니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-03] 무게와 질량을 측정하는 도구의 원리를 이해하고 장소에 따른 측정값의 변화를 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 무게 측정 기구(용수철저울 → 달에서 1/6)와 질량 비교 기구(양팔저울 → 달에서도 불변)의 메커니즘 차이를 올바르게 변별할 수 있는지 평가하고자 출제하였습니다.
    </div>
  </div>

  <!-- 해설 6 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 6 해설] 정답 ① <span class="sol-badge">난이도: 상</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • 추를 달지 않았을 때(0 N) 길이가 10 cm이므로, <b>용수철의 원래 길이는 10 cm</b>입니다.<br>
      • 추의 무게가 2 N 증가할 때마다 전체 길이가 4 cm씩 늘어나므로, <b>추의 무게 1 N당 늘어난 길이는 2 cm</b>입니다.<br>
      • 물체 A를 매달았을 때 전체 길이가 26 cm이므로, <b>늘어난 길이는 26 - 10 = 16 cm</b>입니다.<br>
      • 1 N당 2 cm 늘어나므로, 물체 A의 무게는 16 ÷ 2 = <b>8 N</b>입니다. (① 참)
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>② (거짓)</b>: 원래 길이는 추를 달지 않은 10 cm입니다.<br>
      • <b>③ (거짓 - ★최빈출 함정)</b>: 추의 무게에 정비례하는 것은 '전체 길이'가 아니라 <b>'늘어난 길이'</b>입니다.<br>
      • <b>④ (거짓)</b>: 26 cm는 전체 길이이며, 늘어난 길이는 16 cm입니다.<br>
      • <b>⑤ (거짓)</b>: 10 N을 매달면 늘어난 길이가 20 cm이므로, 전체 길이는 10 + 20 = 30 cm가 됩니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-04] 탄성력의 크기와 방향을 알고 용수철의 늘어난 길이와 추의 무게 사이의 비례 관계를 해석할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 전체 길이와 늘어난 길이를 혼동하는 전형적인 실수를 바로잡고, 원래 길이를 차감하여 비례식을 수립하는 능력을 기르기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 7 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 7 해설] 정답 ③ <span class="sol-badge">난이도: 중</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: (가)와 (나)는 면적만 다르고 무게와 거칠기가 같은데 저울 눈금이 2 N으로 동일하므로, <b>접촉 면적은 마찰력에 영향을 주지 않음</b>을 증명합니다.<br>
      • <b>ㄴ (참)</b>: (가)와 (다)는 같은 나무판 위에서 실험하였으므로, 통제 변인은 '접촉면의 거칠기'이고 조작 변인은 '나무도막의 무게'입니다.<br>
      • <b>ㄷ (참)</b>: (가)와 (라)는 무게가 같고 거칠기만 다른데 사포 위에서 눈금이 5 N으로 증가하였으므로, <b>접촉면이 거칠수록 마찰력이 커짐</b>을 알 수 있습니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • ㄱ, ㄴ, ㄷ 모두 탐구 변인 통제 원리에 정확히 부합하므로 정답은 ③입니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-04] 마찰력의 크기에 영향을 미치는 요인을 탐구 실험을 통해 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 교과서 대표 탐구 실험을 통해 조작 변인과 통제 변인을 체계적으로 분석하고, 마찰력 결정 요인을 검증할 수 있는지 확인하기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 8 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 8 해설] 정답 ⑤ <span class="sol-badge">난이도: 상</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: 상자가 '일정한 속력으로 곧게(등속 직선 운동)' 이동하므로 운동 상태가 변하지 않아 <b>알짜힘은 0 N</b>입니다.<br>
      • <b>ㄴ (참)</b>: 오른쪽으로 6 N을 당길 때 알짜힘이 0 N이 되려면 반대 방향(왼쪽)으로 크기가 같은 <b>6 N의 마찰력</b>이 작용해야 합니다.<br>
      • <b>ㄷ (참)</b>: 벽돌을 올리면 바닥을 누르는 무게가 증가하므로 마찰력이 커져 일정한 속력으로 끌기 위한 힘도 6 N보다 커집니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>ㄹ (거짓 - ★대표 오개념)</b>: 마찰력의 크기는 <b>접촉 면적과 무관</b>합니다. 상자를 세워 접촉 면적을 줄이더라도 상자의 무게와 바닥의 거칠기가 일정하므로 마찰력은 여전히 6 N이며, 필요한 힘도 6 N으로 동일합니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-04] 마찰력의 방향과 크기 결정 요인을 이해하고 일상생활의 운동에 적용할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 등속 운동 시 '외력 = 마찰력'의 평형 관계를 인지하고, 접촉 면적이 줄어들면 마찰력이 감소한다는 상식을 과학적으로 바로잡기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 9 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 9 해설] 정답 ② <span class="sol-badge">난이도: 상</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: 공기 중 10 N인데 (나)에서 8 N이 되었으므로, 부력은 10 - 8 = <b>2 N</b>입니다.<br>
      • <b>ㄴ (참)</b>: 부력은 <b>잠긴 부피에 비례</b>합니다. 절반 잠겼을 때 2 N이므로, 전체가 완전히 잠긴 (다)에서는 부력이 4 N이 됩니다. 따라서 눈금 A = 10 - 4 = <b>6 N</b>입니다.<br>
      • <b>ㄹ (참)</b>: 완전히 잠긴 후에는 깊이가 더 깊어지더라도 잠긴 부피가 같으므로 <b>부력의 크기는 변하지 않습니다.</b>
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>ㄷ (거짓 - ★최빈출 함정)</b>: 완전히 잠긴 후에는 수심이 깊어져도 부력이 4 N으로 일정하므로 저울 눈금 B는 A와 똑같이 6 N입니다. (B &lt; A가 아님)
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-04] 부력의 개념을 이해하고 물체의 잠긴 부피에 따른 부력의 변화를 정량적으로 해석할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 부력은 '잠긴 부피'에 의해서만 결정되며, 물속 깊이(수심)와는 전혀 무관하다는 점을 확실히 숙지시키고자 출제하였습니다.
    </div>
  </div>

  <!-- 해설 10 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 10 해설] 정답 ① <span class="sol-badge">난이도: 중</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>(가) 스카이다이버 낙하</b>: 운동 방향으로 중력이 작용하여 아래로 점점 빨라지므로 <b>속력만 변하는 운동</b>입니다.<br>
      • <b>(나) 대관람차 회전</b>: 일정한 속력으로 회전하므로 <b>운동 방향만 변하는 운동</b>입니다.<br>
      • <b>(다) 비스듬히 찬 축구공</b>: 포물선을 그리며 날아가므로 <b>속력과 운동 방향이 모두 변하는 운동</b>입니다.<br>
      • <b>(라) 수평 컨베이어 벨트</b>: 일정한 속력으로 직선 운동하므로 <b>속력과 방향이 모두 일정한 운동</b>입니다.<br>
      따라서 올바르게 연결된 것은 ①입니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • ②, ③, ④, ⑤는 각 운동 형태의 분류 항목이 뒤바뀌어 오답입니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-05] 힘이 작용할 때 물체의 운동 상태(속력, 운동 방향)가 변하는 다양한 사례를 비교·분류할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 힘의 작용 방향에 따른 4가지 대표 운동 유형을 실생활 상황과 정확히 매칭할 수 있는지 평가하기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 11 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 11 해설] 정답 ⑤ <span class="sol-badge">난이도: 상</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>⑤ (참)</b>: 물체에 작용하는 알짜힘이 0 N일 때, 정지해 있던 물체는 정지 상태를 유지하고, 이미 운동하던 물체는 <b>일정한 속력으로 곧게 운동(등속 직선 운동)</b>을 유지합니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>① (거짓)</b>: 마찰이 없으므로 힘을 가하지 않아도 속력이 일정하게 유지됩니다.<br>
      • <b>② (거짓)</b>: 등속 직선 운동을 하고 있으므로 알짜힘은 0 N입니다.<br>
      • <b>③ (거짓)</b>: 반대 방향으로 힘을 받으면 즉시 방향을 바꾸는 것이 아니라, 서서히 속력이 줄어 정지한 후 반대로 이동합니다.<br>
      • <b>④ (거짓)</b>: 운동 방향과 수직인 힘은 속력은 일정하고 운동 방향만 바꿉니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-05] 힘이 작용하지 않거나 알짜힘이 0일 때 물체의 운동 상태를 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: "운동하려면 힘이 필요하다"는 직관적 오개념을 깨고, 알짜힘이 0일 때 등속 직선 운동을 한다는 핵심 원리를 정립하기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 12 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 12 해설] 정답 ② <span class="sol-badge">난이도: 상</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄴ (참)</b>: 일정한 시간 간격 동안 이동한 구간 거리가 점점 넓어지므로 속력이 일정하게 빨라집니다.<br>
      • <b>ㄷ (참)</b>: 공기 저항이 없는 진공에서는 질량에 관계없이 속력 변화율이 같으므로 동시에 바닥에 도달합니다.<br>
      • <b>ㄹ (참)</b>: 낙하하는 동안 운동 방향과 같은 연직 아래쪽으로 중력이 계속 작용합니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>ㄱ (거짓 - ★중1 최고 함정)</b>: 중력의 크기(무게)는 질량에 비례합니다. 질량이 2 kg인 B에 작용하는 중력(19.6 N)은 질량이 1 kg인 A에 작용하는 중력(9.8 N)의 <b>2배</b>입니다. (동시에 떨어지는 것과 받는 중력의 크기는 별개입니다.)
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과05-06] 자유 낙하 운동을 관찰하고 다중 섬광 사진을 분석하여 속력 변화의 규칙성을 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 동시 도달 현상 때문에 중력의 크기마저 같다고 착각하는 중1 학생들의 대표적 오류를 정확히 짚어내기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 13 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 13 해설] 정답 ③ <span class="sol-badge">난이도: 하</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>③ (참)</b>: 압력은 단위 면적당 작용하는 힘(압력 = 힘 / 면적)이므로, 눈밭에서 바닥 면적이 넓은 설피를 신으면 <b>접촉 면적이 넓어져 눈에 가해지는 압력이 작아지므로</b> 발이 빠지지 않습니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>① (거짓)</b>: 힘이 일정할 때 면적이 넓어질수록 압력은 작아집니다.<br>
      • <b>② (거짓)</b>: 면적이 일정할 때 작용하는 힘이 작아지면 압력도 작아집니다.<br>
      • <b>④ (거짓)</b>: 칼날을 날카롭게 갈아야 면적이 좁아져 압력이 커지므로 잘 깎입니다.<br>
      • <b>⑤ (거짓)</b>: 좁은 면이 닿을 때 면적이 좁아지므로 압력은 더 커집니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과06-01] 일정한 면적에 작용하는 힘으로 압력을 정의하고, 일상생활에서 압력을 높이거나 낮추는 사례를 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 압력의 기본 정의(P = F / A)와 일상생활 속 적용 원리를 점검하기 위해 출제하였습니다.
    </div>
  </div>

  <!-- 해설 14 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 14 해설] 정답 ④ <span class="sol-badge">난이도: 중</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>④ (거짓 - 정답)</b>: 쇠구슬의 수가 30개인 (나)가 15개인 (가)보다 벽에 충돌하는 횟수가 더 많으므로 <b>(나)가 (가)보다 손바닥에 더 큰 힘(더 큰 압력)이 느껴집니다.</b> 따라서 (가)가 더 큰 힘이 느껴진다는 설명은 옳지 않습니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>①, ②, ③ (참)</b>: 쇠구슬은 기체 입자, 충돌은 기체의 벽면 충돌, 손바닥에 느껴지는 힘은 기체의 압력을 나타내는 타당한 모형입니다.<br>
      • <b>⑤ (참)</b>: 기체 입자는 모든 방향으로 끊임없이 운동하므로 압력도 모든 방향으로 작용합니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과06-02] 기체의 압력이 나타나는 원인을 기체 입자의 충돌과 입자 운동 모형으로 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 교과서 수록 쇠구슬 모형 실험을 통해 기체 압력의 본질(입자 충돌 횟수와 압력의 비례)을 올바르게 이해했는지 평가하고자 출제하였습니다.
    </div>
  </div>

  <!-- 해설 15 -->
  <div class="sol-item">
    <div class="sol-q-title">[문제 15 해설] 정답 ④ <span class="sol-badge">난이도: 상</span></div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 정답 분석:</div>
      • <b>ㄱ (참)</b>: 누름못 30개 위에 풍선을 올리면 누르는 힘이 분산되어 총접촉 면적이 넓어지므로 <b>풍선이 받는 압력이 작아져</b> 터지지 않습니다.<br>
      • <b>ㄴ (참)</b>: 빨판 내부의 공기가 빠져나가 내부 기압이 낮아지면, <b>외부 기압이 빨판을 벽으로 밀어붙여</b> 붙어 있게 됩니다.<br>
      • <b>ㄷ (참)</b>: 높은 산으로 갈수록 대기의 양이 줄어들어 외부 기압이 낮아지므로 과자 봉지 내부 기압에 의해 사방으로 부풀어 오릅니다.
    </div>
    <div class="sol-section">
      <div class="sol-section-title">▶ 오답 피하기:</div>
      • <b>ㄹ (거짓 - ★핵심 오개념)</b>: 기체 입자는 한 방향이 아니라 <b>모든 방향으로 끊임없이 불규칙하게 운동</b>하여 벽에 충돌하므로 압력이 모든 방향으로 고르게 작용하여 둥근 형태를 이룹니다.
    </div>
    <div class="why-box">
      <span class="why-title">[문제를 낸 이유]</span><br>
      • <b>성취기준</b>: [9과06-02] 기압이 모든 방향으로 작용함을 알고, 일상생활 속에서 기체의 압력과 관련된 다양한 현상을 통합적으로 설명할 수 있다.<br>
      • <b>출제 배경 및 오개념 방지</b>: 면적과 압력, 기압 차이, 고도에 따른 기압, 전방위 작용성을 종합하여 실생활 과학 탐구 사고력을 심화 평가하기 위해 출제하였습니다.
    </div>
  </div>

</div><!-- end columns -->

</body>
</html>
"""

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"[1] HTML saved: {html_path} ({os.path.getsize(html_path)} bytes)")

    temp_pdf = os.path.join(base_dir, 'temp_exam_out.pdf')
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)

    cmd = [
        chrome_path,
        '--headless=new',
        '--disable-gpu',
        '--no-pdf-header-footer',
        f'--print-to-pdf={temp_pdf}',
        html_path
    ]
    print("Running Chrome headless command...")
    res = subprocess.run(cmd, capture_output=True)
    if os.path.exists(temp_pdf) and os.path.getsize(temp_pdf) > 1000:
        shutil.move(temp_pdf, pdf_path)
        print(f"[SUCCESS] PDF successfully generated: {pdf_path} ({os.path.getsize(pdf_path)} bytes)")
        doc = fitz.open(pdf_path)
        print(f"[VERIFY] Total pages in generated PDF: {len(doc)}")
        for i in range(len(doc)):
            print(f"  - Page {i+1}: {len(doc[i].get_text())} characters")
    else:
        print(f"[FAIL] Return code: {res.returncode}, Stderr: {res.stderr.decode('utf-8', errors='ignore')}")

if __name__ == '__main__':
    import shutil
    generate_exam_pdf()
