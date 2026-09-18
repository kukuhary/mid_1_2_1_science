import os
import subprocess
import fitz

def build_summary_v2_pdf():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    html_path = os.path.join(base_dir, 'summary_v2.html')
    pdf_path = os.path.join(base_dir, '5단원_힘의_작용_핵심요약_v2.pdf')
    chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

    html_content = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>[2026 중간고사 대비] 중1 과학 핵심 요약: Ⅴ. 힘의 작용 (v2.0 빈출·함정 정밀 분석)</title>
<style>
  @page {
    size: A4;
    margin: 10mm 9mm 10mm 9mm;
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
    line-height: 1.68;
    font-size: 9.5pt;
    margin: 0;
    padding: 0;
    background: #fff;
  }

  /* 1단 전폭 상단 헤더 */
  .header-box {
    border: 2px solid #000;
    padding: 7px 12px;
    margin-bottom: 9px;
    text-align: center;
    background-color: #fff;
  }
  .header-title {
    font-size: 14pt;
    font-weight: bold;
    margin: 0 0 4px 0;
    letter-spacing: -0.5px;
  }
  .header-info-table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 3px;
    font-size: 8.5pt;
  }
  .header-info-table td {
    border: 1px solid #444;
    padding: 3px 5px;
    text-align: center;
    background-color: #f7f7f7;
  }
  .header-info-table td.label {
    font-weight: bold;
    background-color: #eaeaea;
    width: 13%;
  }

  /* 2단 레이아웃 */
  .two-column-layout {
    column-count: 2;
    column-gap: 7mm;
    column-rule: 0.8px solid #777;
    text-align: justify;
  }

  /* 섹션 구분 바 */
  .section-bar {
    font-size: 10pt;
    font-weight: bold;
    border-top: 1.5px solid #000;
    border-bottom: 1px solid #000;
    padding: 3px 5px;
    margin: 8px 0 6px 0;
    background-color: #ededed;
    break-after: avoid;
    -webkit-column-break-after: avoid;
  }

  /* 소주제 헤더 */
  .sub-topic {
    font-size: 9.5pt;
    font-weight: bold;
    margin: 6px 0 3px 0;
    border-left: 3px solid #000;
    padding-left: 5px;
    break-after: avoid;
    -webkit-column-break-after: avoid;
  }

  /* 일반 개념 박스 */
  .concept-box {
    border: 1px solid #bbb;
    background-color: #fafafa;
    padding: 5px 7px;
    margin-bottom: 6px;
    font-size: 9.0pt;
    line-height: 1.6;
    break-inside: avoid;
    -webkit-column-break-inside: avoid;
  }

  /* 2-Depth 목차 전용 박스 */
  .toc-box {
    border: 1.2px solid #222;
    background-color: #fcfcfc;
    padding: 6px 8px;
    margin-bottom: 8px;
    font-size: 8.6pt;
    line-height: 1.55;
    break-inside: avoid;
    -webkit-column-break-inside: avoid;
  }
  .toc-header {
    font-weight: bold;
    font-size: 9.2pt;
    text-align: center;
    border-bottom: 1px solid #666;
    padding-bottom: 2px;
    margin-bottom: 5px;
    background-color: #efefef;
  }
  .toc-depth1 {
    font-weight: bold;
    color: #000;
    margin-top: 3px;
  }
  .toc-depth2 {
    margin-left: 6px;
    color: #111;
  }
  .toc-desc {
    color: #444;
    margin-left: 12px;
    font-size: 8.2pt;
  }

  /* [빈출] 태그 및 스타일 */
  .tag-frequent {
    display: inline-block;
    font-size: 8.2pt;
    font-weight: bold;
    border: 1px solid #444;
    background-color: #e6e6e6;
    padding: 0 4px;
    border-radius: 2px;
    margin-right: 3px;
  }

  /* [함정 ★최빈출] 박스 - 함정 최우선 강조 */
  .trap-box {
    border: 1.4px solid #000;
    background-color: #f2f2f2;
    padding: 6px 8px;
    margin: 5px 0 7px 0;
    font-size: 8.9pt;
    line-height: 1.58;
    break-inside: avoid;
    -webkit-column-break-inside: avoid;
  }
  .trap-tag {
    display: inline-block;
    font-size: 8.2pt;
    font-weight: bold;
    background-color: #000;
    color: #fff;
    padding: 1px 5px;
    margin-right: 4px;
  }
  .trap-content {
    margin-top: 3px;
  }

  ul {
    margin: 2px 0 4px 0;
    padding-left: 15px;
  }
  li {
    margin-bottom: 2px;
  }
  .formula {
    font-family: 'Times New Roman', serif;
    font-style: italic;
    font-weight: bold;
  }
</style>
</head>
<body>

  <!-- 1단 전폭 상단 헤더 -->
  <div class="header-box">
    <div class="header-title">[2026학년도 중간고사 대비] 중학교 1학년 과학 핵심 요약 (v2.0)</div>
    <table class="header-info-table">
      <tr>
        <td class="label">과목 및 학년</td>
        <td>중학교 과학 1 (중1)</td>
        <td class="label">단원 영역</td>
        <td><b>Ⅴ. 힘의 작용 (154~193쪽)</b></td>
        <td class="label">교과서 출처</td>
        <td>㈜비상교육 (임태훈 외)</td>
      </tr>
      <tr>
        <td class="label">표준 규격</td>
        <td>2-Depth 직관적 집약 구조</td>
        <td class="label">특화 분석</td>
        <td><b>[빈출] 및 [함정] 정밀 분석 (함정 최우선 원칙)</b></td>
        <td class="label">문서 버전</td>
        <td>v2.0 (실전 시험 대비용)</td>
      </tr>
    </table>
  </div>

  <!-- 본문 2단 조판 시작 -->
  <div class="two-column-layout">

    <!-- 2-Depth 목차 요약 박스 -->
    <div class="toc-box">
      <div class="toc-header">■ Ⅴ. 힘의 작용 2-Depth 핵심 집약 목차</div>
      
      <div class="toc-depth1">[1] 힘의 표현과 평형</div>
      <div class="toc-depth2">• 과학에서의 힘과 표현</div>
      <div class="toc-desc">- 물체의 모양과 운동 상태를 바꾸는 원인 (작용점·방향·크기)</div>
      <div class="toc-depth2">• 합력과 두 힘의 평형</div>
      <div class="toc-desc">- 합력이 0이 될 때 물체는 멈추거나 계속 직진한다</div>

      <div class="toc-depth1">[2] 여러 가지 힘</div>
      <div class="toc-depth2">• 중력과 무게·질량</div>
      <div class="toc-desc">- 지구 중심이 당기는 힘과 변하지 않는 고유한 양</div>
      <div class="toc-depth2">• 탄성력</div>
      <div class="toc-desc">- 변형된 길이에 비례하여 원래대로 되돌아가려는 힘</div>
      <div class="toc-depth2">• 마찰력</div>
      <div class="toc-desc">- 접촉면의 거칠기와 무게가 미끄러짐을 방해한다</div>
      <div class="toc-depth2">• 부력</div>
      <div class="toc-desc">- 액체에 잠긴 부피만큼 물체를 위로 밀어 올린다</div>

      <div class="toc-depth1">[3] 힘의 작용과 운동 상태 변화</div>
      <div class="toc-depth2">• 알짜힘과 운동 상태</div>
      <div class="toc-desc">- 힘이 0이면 유지되고, 힘을 받으면 상태가 변한다</div>
      <div class="toc-depth2">• 힘의 방향과 운동 변화</div>
      <div class="toc-desc">- 나란하면 빠르기가, 수직이면 방향이 바뀐다</div>
      <div class="toc-depth2">• 자유 낙하 운동</div>
      <div class="toc-desc">- 오직 중력만을 받아 속력이 일정하게 빨라진다</div>
    </div>

    <!-- [1] 힘의 표현과 평형 -->
    <div class="section-bar">[1] 힘의 표현과 평형</div>

    <div class="sub-topic">1. 과학에서의 힘과 표현 방법</div>
    <div class="concept-box">
      <ul>
        <li><b>힘의 정의</b>: 물체의 <b>모양</b>을 변하게 하거나, <b>운동 상태(속력·방향)</b>를 변하게 하는 원인 (단위: N)</li>
        <li><span class="tag-frequent">빈출</span> <b>힘의 3요소와 화살표</b>:
          <ul>
            <li>힘의 3요소: <b>작용점, 방향, 크기</b></li>
            <li>시작점: 작용점 / 머리: 힘의 방향 / 길이: 힘의 크기에 비례</li>
          </ul>
        </li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 주의</span> <b>일상어 착각 & 작용점 위치</b>
      <div class="trap-content">
        ① "책을 읽어 힘이 들었다", "설득하는 힘" 등 피로감이나 능력은 과학에서의 힘이 아님!<br>
        ② 힘의 크기와 방향이 같더라도 <b>작용점 위치가 다르면 회전 등 물체에 미치는 효과가 완전히 달라짐</b>.
      </div>
    </div>

    <div class="sub-topic">2. 나란한 두 힘의 합력과 평형</div>
    <div class="concept-box">
      <ul>
        <li><span class="tag-frequent">빈출</span> <b>나란한 두 힘의 합력</b>:
          <ul>
            <li>같은 방향: <span class="formula">F = F₁ + F₂</span> (방향: 두 힘의 방향)</li>
            <li>반대 방향: <span class="formula">F = F_큰 - F_작은</span> (방향: 더 큰 힘의 방향)</li>
          </ul>
        </li>
        <li><b>두 힘의 평형 3대 조건</b>: ① 크기 같음, ② 방향 정반대, ③ 동일 작용선상.</li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 ★최빈출</span> <b>작용점 한 물체 & 등속 운동 착각</b>
      <div class="trap-content">
        ① <b>반드시 '한 물체'에 동시 작용</b>해야 평형 성립! 서로 다른 두 물체에 작용하는 힘(작용·반작용)은 결코 평형이 아님.<br>
        ② <b>등속 직선 운동</b>을 하는 물체는 힘을 계속 받는 것이 아니라 <b>알짜힘 = 0 N인 힘의 평형 상태</b>임!
      </div>
    </div>

    <!-- [2] 여러 가지 힘 -->
    <div class="section-bar">[2] 여러 가지 힘</div>

    <div class="sub-topic">1. 중력과 무게·질량</div>
    <div class="concept-box">
      <ul>
        <li><b>중력</b>: 지구가 물체를 지구 중심(연직 아래)으로 당기는 힘 (질량에 비례)</li>
        <li><span class="tag-frequent">빈출</span> <b>무게 계산</b>: 지구 표면 질량 1 kg ≈ 9.8 N / 달의 중력은 지구의 약 1/6</li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 ★최빈출</span> <b>달에서의 질량 불변 & 저울 구별</b>
      <div class="trap-content">
        ① 달에 가면 무게(N)만 1/6로 줄고, 물질의 고유한 양인 <b>질량(kg)은 절대 불변</b>!<br>
        ② <b>윗접시·양팔저울</b>은 분동과 비교하므로 달에서도 <b>눈금 불변</b>! 오직 <b>용수철·앉은뱅이저울</b>만 무게를 측정해 1/6로 감소.
      </div>
    </div>

    <div class="sub-topic">2. 탄성력</div>
    <div class="concept-box">
      <ul>
        <li><b>복원 방향</b>: 변형된 방향의 정반대 (가한 외력의 반대 방향)</li>
        <li><span class="tag-frequent">빈출</span> <b>이용</b>: 용수철저울, 볼펜 스프링, 트램펄린, 활, 장대높이뛰기</li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 ★최빈출</span> <b>'늘어난 길이' vs '전체 길이'</b>
      <div class="trap-content">
        추 무게에 정비례하는 것은 전체 길이가 아닌 <b>'늘어난 길이(변형 길이)'</b>임! 그래프 분석 시 0 N일 때의 원래 길이(y절편)를 전체 길이에서 반드시 빼야 함.<br>
        <i>(늘어난 길이 = 전체 길이 - 원래 길이)</i>
      </div>
    </div>

    <div class="sub-topic">3. 마찰력</div>
    <div class="concept-box">
      <ul>
        <li><b>방해 방향</b>: 물체가 미끄러지려는 방향(운동 방향)의 정반대</li>
        <li><span class="tag-frequent">빈출</span> <b>결정 요인</b>: 접촉면이 거칠수록 큼, 물체가 무거울수록(수직 누름 클수록) 큼</li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 ★최빈출</span> <b>접촉 면적 무관성 & 등속 당김</b>
      <div class="trap-content">
        ① <b>접촉 면적의 넓이와 마찰력은 완전 무관</b>함! (도막을 넓게 눕히든 좁게 세우든 마찰력 동일).<br>
        ② 수평면에서 일정한 빠르기(등속 운동)로 끌어당길 때, 알짜힘이 0 N이므로 <b>용수철저울 눈금(끄는 힘) = 마찰력</b>이 됨.
      </div>
    </div>

    <div class="sub-topic">4. 부력</div>
    <div class="concept-box">
      <ul>
        <li><b>방향</b>: 액체나 기체 속에서 물체를 위로 밀어 올리는 힘 (연직 위쪽, 중력 반대)</li>
        <li><span class="tag-frequent">빈출</span> <b>크기 및 수중 무게</b>: 잠긴 부피에 정비례 / 수중 무게 = 공기 중 무게 - 부력</li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 ★최빈출</span> <b>수심 무관성 & 떠 있는 물체</b>
      <div class="trap-content">
        ① 물에 완전히 잠긴 후에는 <b>아무리 깊이 들어가도 부력은 변하지 않고 항상 일정</b>함! (부력은 수심과 무관, 오직 잠긴 부피에만 비례).<br>
        ② 수면에 정지해 떠 있는 나무도막은 부력을 안 받는 게 아니라, <b>'부력 = 중력'으로 평형</b>을 이루어 떠 있는 것임.
      </div>
    </div>

    <!-- [3] 힘의 작용과 운동 상태 변화 -->
    <div class="section-bar">[3] 힘의 작용과 운동 상태 변화</div>

    <div class="sub-topic">1. 알짜힘과 운동 상태</div>
    <div class="concept-box">
      <ul>
        <li><b>운동 상태</b>: 물체의 <b>빠르기(속력)</b>와 <b>운동 방향</b></li>
        <li><span class="tag-frequent">빈출</span> <b>알짜힘 유무</b>: 알짜힘 = 0 (정지 유지/등속 운동) vs 알짜힘 ≠ 0 (상태 변함)</li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 ★최빈출</span> <b>운동 유지와 힘의 관계 착각</b>
      <div class="trap-content">
        "움직이려면 힘이 계속 작용해야 한다?" $\rightarrow$ <b>가장 흔한 오답!</b> 힘은 운동을 '유지'하는 원인이 아니라 '변화'시키는 원인임. 힘이 0 N이어도 물체는 등속 직선 운동을 영원히 지속함.
      </div>
    </div>

    <div class="sub-topic">2. 힘의 방향과 운동 변화</div>
    <div class="concept-box">
      <ul>
        <li><span class="tag-frequent">빈출</span> <b>힘의 방향에 따른 운동 분류</b>:
          <ul>
            <li>나란한 힘: 속력만 변함 (같은 방향: 빨라짐 / 반대 방향: 느려짐)</li>
            <li>수직인 힘: 빠르기 일정, 운동 방향만 계속 바뀜</li>
            <li>비스듬한 힘: 빠르기와 운동 방향이 동시에 바뀜 (그네, 바이킹)</li>
          </ul>
        </li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 주의</span> <b>등속 원운동의 알짜힘 존재 여부</b>
      <div class="trap-content">
        회전목마, 대관람차처럼 일정한 속력으로 도는 운동은 속력은 일정해도 <b>운동 방향이 매 순간 바뀌므로 알짜힘이 0이 아니며, 원 중심 방향으로 힘을 계속 받는 운동</b>임!
      </div>
    </div>

    <div class="sub-topic">3. 자유 낙하 운동과 다중 섬광 사진</div>
    <div class="concept-box">
      <ul>
        <li><b>본질</b>: 공기 저항 없이 오직 중력만 받아 아래로 떨어지는 운동</li>
        <li><span class="tag-frequent">빈출</span> <b>특징 & 다중 섬광 사진</b>: 운동 방향으로 일정한 중력을 계속 받아 매초 속력 일정 증가, 구간 간격이 아래로 갈수록 점점 넓어짐.</li>
      </ul>
    </div>
    <div class="trap-box">
      <span class="trap-tag">함정 ★최빈출</span> <b>진공 낙하 질량 무관 & 낙하 중 힘</b>
      <div class="trap-content">
        ① 진공에서는 쇠구슬과 깃털이 <b>질량과 무관하게 동시에 낙하</b>함! (속력 증가율은 질량과 무관).<br>
        ② 낙하 중 속력이 빨라진다고 중력이 커지는 것이 아니며, <b>항상 일정한 크기의 중력(알짜힘)</b>을 받음.
      </div>
    </div>

  </div>
  <!-- 본문 2단 조판 끝 -->

</body>
</html>
"""

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"[1] HTML v2 created: {html_path} ({os.path.getsize(html_path)} bytes)")

    # Execute Chrome headless to print PDF
    cmd = [
        chrome_path,
        '--headless=new',
        '--disable-gpu',
        '--no-pdf-header-footer',
        f'--print-to-pdf={pdf_path}',
        html_path
    ]
    res = subprocess.run(cmd, capture_output=True)
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
        print(f"[2] PDF v2 successfully generated: {pdf_path} ({os.path.getsize(pdf_path)} bytes)")
    else:
        print(f"[FAIL] Return code: {res.returncode}, Stderr: {res.stderr.decode('utf-8', errors='ignore')}")
        return

    # Verify with PyMuPDF
    doc = fitz.open(pdf_path)
    print(f"[3] PDF v2 Page Count: {len(doc)}")
    for i, page in enumerate(doc):
        text = page.get_text()
        print(f"Page {i+1}: Length {len(text)} chars, Preview: {text[:80].strip()}...")

if __name__ == '__main__':
    build_summary_v2_pdf()
