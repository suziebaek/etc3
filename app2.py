import streamlit as st
import docx
import re
import html
import io
import zipfile

# ==========================================
# [기능 1] 수능 독해용 Word 파일 정밀 파싱 로직
# ==========================================
def parse_reading_docx(file):
    doc = docx.Document(file)
    questions = []
    current_q = None
    
    # 더 유연한 정규식으로 선지 및 번호 매칭율 상향
    q_pattern = re.compile(r'^(\d+)\.?\s*(.*)')  
    opt_pattern = re.compile(r'^([①②③④⑤\d])\s*(.*)') 
    
    all_elements = []
    
    # 1차 순회: 본문 및 표 요소를 순서대로 정렬 수집
    for element in doc.element.body:
        if element.getparent() != doc.element.body:
            continue
            
        if element.tag.endswith('p'):  # 일반 문단
            p = docx.text.paragraph.Paragraph(element, doc)
            text_with_formatting = ""
            for run in p.runs:
                if run.underline and run.text.strip():
                    text_with_formatting += f"<u>{run.text}</u>"
                else:
                    text_with_formatting += run.text
            
            txt = text_with_formatting.strip()
            if txt:
                all_elements.append({"type": "text", "text": txt})
                
        elif element.tag.endswith('tbl'):  # 표/상자 요소
            t = docx.table.Table(element, doc)
            table_lines = []
            for row in t.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        cell_txt_formatted = ""
                        for run in paragraph.runs:
                            if run.underline and run.text.strip():
                                cell_txt_formatted += f"<u>{run.text}</u>"
                            else:
                                cell_txt_formatted += run.text
                        ctxt = cell_txt_formatted.strip()
                        if ctxt and ctxt not in table_lines:
                            table_lines.append(ctxt)
            if table_lines:
                all_elements.append({"type": "table", "lines": table_lines})

    # 2차 순회: 수정된 번호별 분기 조건 가공
    for item in all_elements:
        if item["type"] == "table":
            if current_q is not None:
                q_num_int = int(current_q["num"]) if current_q["num"].isdigit() else 1
                for t_line in item["lines"]:
                    # 밑줄 서식 교정 규칙 적용
                    converted_t_line = re.sub(r'_{2,}', '<span class="underline" style="width: 120px; display: inline-block; border-bottom: 1px solid #000;"></span>', t_line)
                    
                    if 19 <= q_num_int <= 22:
                        current_q["box_sentence"].append(converted_t_line)
                    else:
                        current_q["passage_lines"].append(converted_t_line)
            continue

        line = item["text"].strip()

        # 새로운 문제 번호 등장 감지
        q_match = q_pattern.match(line)
        if q_match:
            if current_q is not None:
                questions.append(current_q)
            current_q = {
                "num": q_match.group(1),
                "title": q_match.group(2),
                "passage_lines": [], 
                "box_sentence": [],  
                "options": []
            }
            continue

        if current_q is None:
            continue

        if line.startswith("정답:") or "→" in line:
            continue
        if line.startswith("[") and "]" in line:
            continue

        # 선택지 보기(①~⑤) 매칭 처리 (선지 유실 방지 디펜스)
        opt_match = opt_pattern.match(line)
        if opt_match and any(char in line[:3] for char in ["①", "②", "③", "④", "⑤"]):
            # 맨 앞 글자 추출 및 숫자 매핑
            char = line[0]
            label_map = {"①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5"}
            num = label_map.get(char, "1")
            opt_text = line[1:].strip()
            
            processed_opt = re.sub(r'\s{2,}', ' &nbsp;&nbsp;&nbsp;&nbsp; ', opt_text)
            current_q["options"].append({
                "char": char,
                "num": num,
                "text": processed_opt
            })
            continue

        # 일반 지문 영역 및 밑줄 태그 치환 규칙 적용
        converted_line = re.sub(r'_{2,}', '<span class="underline" style="width: 120px; display: inline-block; border-bottom: 1px solid #000;"></span>', line)
        current_q["passage_lines"].append(converted_line)

    if current_q is not None:
        questions.append(current_q)

    return questions


# ==========================================
# [기능 2] 번호대별 스킨 분기형 HTML 렌더러
# ==========================================
def generate_split_reading_html(q):
    q_num_int = int(q["num"]) if q["num"].isdigit() else 1

    # 공통 상단 헤더
    html_content = """<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta charset="UTF-8" name="viewport" content="width=device-width, target-densitydpi=device-dpi" />
<meta HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no,minimum-scale=1.0,maximum-scale=1.0,target-densitydpi=device-dpi">
<title>Achievement Test</title>
<link rel="stylesheet" href="../../../common/css/common.css">
<link rel="stylesheet" href="../../../common/css/style.css">
<link rel="stylesheet" href="../../../common_ex/css/style.css">
<link rel="stylesheet" href="../../../common_ex/css/scroll.css">
<script type="text/javascript" charset="UTF-8" src="../../../common/js/jquery.js"></script>
<script type="text/javascript" charset="UTF-8" src="../../../common/js/common.js"></script>
<script type="text/javascript" charset="UTF-8" src="../../../common_ex/js/common.js"></script>
</head>
<body>
\t<div class="pageWrap">
\t\t<div id="R_question" class="STSection">
\t\t\t<div class="desc_box"><b>Questions 1-28</b> 지문을 읽고 문제를 풀이하시오.</div>
\t\t\t<div class="pageConts">
"""

    safe_title = html.escape(q['title'])
    safe_title = safe_title.replace("&lt;u&gt;", "<u>").replace("&lt;/u&gt;", "</u>")

    # ------------------------------------------
    # 유형 1: 1번 ~ 18번 구조 (14~17번 포함 스타일 정형화 완료)
    # ------------------------------------------
    if q_num_int <= 18:
        # prompt1 영역 생성
        html_content += "\t\t\t\t<div class=\"prompt1\">\n\t\t\t\t\t<div class=\"passage\">\n"
        for p_line in q["passage_lines"]:
            if p_line.startswith("Dear") or p_line.startswith("Sincerely"):
                html_content += f"\t\t\t\t\t\t<p class=\"dialog\">{p_line}</p>\n"
            elif p_line.startswith("*"):
                html_content += f"\t\t\t\t\t\t<p class=\"footnote right\">{p_line}</p>\n"
            else:
                html_content += f"\t\t\t\t\t\t<p class=\"indent\">{p_line}</p>\n"
        html_content += "\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

        # prompt2 영역 생성
        html_content += f"""\t\t\t\t<div class="prompt2">
\t\t\t\t\t<div class="q_box">
\t\t\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer=\"3\" gravity=\"top|left\">
\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t\t\t<td>{safe_title} <br></td>
\t\t\t\t\t\t\t</tr>\n"""
        
        # 선지 마크업 주입
        for opt in q["options"]:
            html_content += f"""\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t\t\t<td><span class="STChoice" remarkable="true" label="{opt['num']}"><span class="label">{opt['char']}</span> {opt['text']}</span></td>
\t\t\t\t\t\t\t</tr>\n"""
        html_content += "\t\t\t\t\t\t</table>\n\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

    # ------------------------------------------
    # 유형 2: 19번 ~ 22번 구조 (colspan="5" 및 sentence 적용 구역)
    # ------------------------------------------
    elif 19 <= q_num_int <= 22:
        # prompt1 구역 처리
        html_content += "\t\t\t\t<div class=\"prompt1\">\n\t\t\t\t\t<div class=\"passage\">\n"
        for p_line in q["passage_lines"]:
            if p_line.startswith("*"):
                html_content += f"\t\t\t\t\t\t<p class=\"footnote right\">{p_line}</p>\n"
            else:
                html_content += f"\t\t\t\t\t\t<p class=\"indent\">{p_line}</p>\n"
        html_content += "\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

        # prompt2 내부 구성
        html_content += f"""\t\t\t\t<div class="prompt2">
\t\t\t\t\t<div class="q_box">
\t\t\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer="4" gravity="top|left">
\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t\t\t<td>{safe_title}<br></td>
\t\t\t\t\t\t\t</tr>\n"""
        
        if q["box_sentence"]:
            box_join = " ".join(q["box_sentence"])
            html_content += f"""\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td colspan="5">
\t\t\t\t\t\t\t\t\t<div class="sentence">{box_join}</div>
\t\t\t\t\t\t\t\t</td>
\t\t\t\t\t\t\t</tr>\n"""

        for opt in q["options"]:
            html_content += f"""\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t\t\t<td><span class="STChoice" remarkable="true" label="{opt['num']}"><span class="label">{opt['char']}</span> {opt['text']}</span></td>
\t\t\t\t\t\t\t</tr>\n"""
        html_content += "\t\t\t\t\t\t</table>\n\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

    # ------------------------------------------
    # 유형 3: 23번 ~ 마지막 번호 구조 (영어 지문 강제 인식 및 문단별 indent 구성)
    # ------------------------------------------
    else:
        # prompt1 : 모든 영어 지문을 문단별로 <p class="indent"> 래핑
        html_content += "\t\t\t\t<div class=\"prompt1\">\n\t\t\t\t\t<div class=\"passage\">\n"
        for p_line in q["passage_lines"]:
            if p_line.startswith("*"):
                html_content += f"\t\t\t\t\t\t<p class=\"footnote right\">{p_line}</p>\n"
            else:
                html_content += f"\t\t\t\t\t\t<p class=\"indent\">{p_line}</p>\n"
        html_content += "\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

        # prompt2 빌드
        html_content += f"""\t\t\t\t<div class="prompt2">
\t\t\t\t\t<div class="q_box">
\t\t\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer="4" gravity="top|left">
\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t\t\t<td>{safe_title}<br></td>
\t\t\t\t\t\t\t</tr>\n"""

        for opt in q["options"]:
            html_content += f"""\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t\t\t<td><span class="STChoice" remarkable="true" label="{opt['num']}"><span class="label">{opt['char']}</span> {opt['text']}</span></td>
\t\t\t\t\t\t\t</tr>\n"""
        html_content += "\t\t\t\t\t\t</table>\n\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

    # 공통 푸터 하단 마감
    html_content += """\t\t\t</div>
\t\t</div>
\t</div>
</body>
</html>"""

    return html_content


# ==========================================
# [기능 3] Streamlit 레이아웃 연동부
# ==========================================
st.set_page_config(page_title="수능 독해형 완전 자동화 패키저", layout="wide")

st.title("🚀 수능형 독해 마크업 정밀 교정 시스템")
st.caption("요청하신 밑줄 서식 매칭, 선지 누락 디펜스, 14~17번 스킨 통일, 19번 이후 문단별 indent 구현이 적용된 버전입니다.")

uploaded_file = st.file_uploader("수능 독해형 워드 문서(.docx)를 업로드하세요", type=["docx"])
submit_button = st.button("🎁 수정사항 반영된 압축 패키지 일괄 생성", type="primary")

if uploaded_file is not None and submit_button:
    try:
        with st.spinner("요청 피드백 반영 및 압축 패키징 중..."):
            parsed_data = parse_reading_docx(uploaded_file)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for q in parsed_data:
                    q_num = q["num"]
                    if not q_num:
                        continue
                    
                    single_html = generate_split_reading_html(q)
                    folder_file_path = f"{q_num}/test.html"
                    zip_file.writestr(folder_file_path, single_html)
            
            zip_data = zip_buffer.getvalue()
            
        st.success(f"🎉 성공! 총 {len(parsed_data)}문항의 수능 독해 전용 완결형 패키지가 완성되었습니다.")
        
        st.subheader("📂 완료된 ZIP 파일 다운로드")
        st.download_button(
            label="📥 피드백 반영 통합 ZIP 다운로드",
            data=zip_data,
            file_name="perfect_reading_test_package.zip",
            mime="application/zip"
        )
            
    except Exception as e:
        st.error(f"⚠️ 처리 중 알 수 없는 예외 오류 발생: {e}")
