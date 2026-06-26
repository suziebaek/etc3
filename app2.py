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
    
    q_pattern = re.compile(r'^(\d+)\.?\s*(.*)')  # 문제 번호 (예: 1. 다음 글의)
    opt_pattern = re.compile(r'^([①②③④⑤])\s*(.*)') # 선택지 (예: ① 대회 일정을)
    
    all_elements = []
    
    # 1차 순회: 최상위 body 기준 문단과 표(지문 박스)를 순서대로 수집
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
                
        elif element.tag.endswith('tbl'):  # 📦 표 (지문 상자)
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

    # 2차 순회: 번호별 수능형 독해 레이아웃 매칭 가공
    for item in all_elements:
        # [A] 박스형 지문(표)을 만났을 때 처리
        if item["type"] == "table":
            if current_q is not None:
                q_num_int = int(current_q["num"]) if current_q["num"].isdigit() else 1
                for t_line in item["lines"]:
                    converted_t_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', t_line)
                    
                    # 💡 [요청 반영] 19~22번까지만 sentence 박스로 수집하고, 23번부터는 다시 일반 지문(passage)으로 전환
                    if 19 <= q_num_int <= 22:
                        current_q["box_sentence"].append(converted_t_line)
                    else:
                        current_q["passage_lines"].append(converted_t_line)
            continue

        line = item["text"].strip()

        # [B] 새로운 문제 번호 등장 시 마감 및 새 구조 방 개설
        q_match = q_pattern.match(line)
        if q_match:
            if current_q is not None:
                questions.append(current_q)
            current_q = {
                "num": q_match.group(1),
                "title": q_match.group(2),
                "passage_lines": [], # prompt1 구역에 위치할 상단 지문들
                "box_sentence": [],  # prompt2 테이블 내부에 들어갈 상자 제시문
                "options": []
            }
            continue

        if current_q is None:
            continue

        # 교사용 정답 및 해설 라인 패스
        if line.startswith("정답:") or "→" in line:
            continue
            
        # 대괄호 지문 안내선 패스 (예: [26~28])
        if line.startswith("[") and "]" in line:
            continue

        # [C] 선택지 보기(①~⑤)를 만났을 때
        opt_match = opt_pattern.match(line)
        if opt_match:
            label_char = opt_match.group(1)
            opt_text = opt_match.group(2)
            
            label_map = {"①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5"}
            label_num = label_map.get(label_char, "1")
            
            processed_opt = re.sub(r'\s{2,}', ' &nbsp;&nbsp;&nbsp;&nbsp; ', opt_text)
            current_q["options"].append({
                "char": label_char,
                "num": label_num,
                "text": processed_opt
            })
            continue

        # [D] 선지 이전의 모든 지문이나 각주 서식 처리
        converted_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', line)
        current_q["passage_lines"].append(converted_line)

    if current_q is not None:
        questions.append(current_q)

    return questions


# ==========================================
# [기능 2] 번호대별 스킨 분기형 HTML 렌더러
# ==========================================
def generate_split_reading_html(q):
    q_num_int = int(q["num"]) if q["num"].isdigit() else 1

    # 공통 헤더 템플릿
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
\t\t<div class="pageWrap">

\t\t\t\t<div id="R_question" class="STSection">
\t\t\t<div class="desc_box"><b>Questions 1-28</b> 지문을 읽고 문제를 풀이하시오.</div>
 
\t\t\t<div class="pageConts">
"""

    # 지시문 내 밑줄 태그 치환 안전성 복원
    safe_title = html.escape(q['title'])
    safe_title = safe_title.replace("&lt;u&gt;", "<u>").replace("&lt;/u&gt;", "</u>")

    # ------------------------------------------
    # 스킨 1: 1번 ~ 18번 유형 레이아웃 (기존 test.html 형태)
    # ------------------------------------------
    if q_num_int <= 18:
        html_content += "\t\t\t\t<div class=\"prompt1\">\n\t\t\t\t\t<div class=\"passage\">\n"
        for p_line in q["passage_lines"]:
            if p_line.startswith("Dear") or p_line.startswith("Sincerely"):
                html_content += f"\t\t\t\t\t\t<p class=\"dialog\">{p_line}</p>\n"
            elif p_line.startswith("*"):
                html_content += f"\t\t\t\t\t\t<p class=\"footnote right\">{p_line}</p>\n"
            else:
                html_content += f"\t\t\t\t\t\t<p class=\"indent\">{p_line}</p>\n"
        html_content += "\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

        html_content += f"""\t\t\t\t<div class="prompt2">
\t\t\t\t\t<div class="q_box">
\t\t\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer="3" gravity="top|left">
\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t\t\t<td>{safe_title} <br></td>
\t\t\t\t\t\t\t</tr>\n"""
        
        for opt in q["options"]:
            html_content += f"""\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t\t\t<td><span class="STChoice" remarkable="true" label="{opt['num']}"><span class="label">{opt['char']}</span> {opt['text']}</span></td>
\t\t\t\t\t\t\t</tr>\n"""
            
        html_content += "\t\t\t\t\t\t</table>\n\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

    # ------------------------------------------
    # 스킨 2: 19번 ~ 22번 유형 레이아웃 (sentence 확장 및 아래 지문 indent 적용)
    # ------------------------------------------
    elif 19 <= q_num_int <= 22:
        # prompt1 : 네모 상자용 제시문 밑에 있는 지문은 전부 <p class="indent"> 처리
        html_content += "\t\t\t\t<div class=\"prompt1\">\n\t\t\t\t\t<div class=\"passage\">\n"
        for p_line in q["passage_lines"]:
            if p_line.startswith("*"):
                html_content += f"\t\t\t\t\t\t<p class=\"footnote right\">{p_line}</p>\n"
            else:
                html_content += f"\t\t\t\t\t\t<p class=\"indent\">{p_line}</p>\n"
        html_content += "\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

        # prompt2 : 문항 내부 테이블 선언
        html_content += f"""\t\t\t\t<div class="prompt2">
\t\t\t\t\t<div class="q_box">
\t\t\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer="4" gravity="top|left">
\t\t\t\t\t\t\t<tr>
\t\t\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t\t\t<td>{safe_title}<br></td>
\t\t\t\t\t\t\t</tr>\n"""
        
        # 💡 [요청 반영] 19번부터 네모 상자용 제시문 구조를 colspan="5"로 감싸기
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
    # 스킨 3: 23번 이후 유형 레이아웃 (네모 상자 지문도 다시전부 indent 처리)
    # ------------------------------------------
    else:
        # prompt1 : 박스 내부에 있던 문장들까지 흡수하여 전부 <p class="indent"> 처리
        html_content += "\t\t\t\t<div class=\"prompt1\">\n\t\t\t\t\t<div class=\"passage\">\n"
        for p_line in q["passage_lines"]:
            if p_line.startswith("*"):
                html_content += f"\t\t\t\t\t\t<p class=\"footnote right\">{p_line}</p>\n"
            else:
                html_content += f"\t\t\t\t\t\t<p class=\"indent\">{p_line}</p>\n"
        html_content += "\t\t\t\t\t</div>\n\t\t\t\t</div>\n"

        # prompt2 : 상자 지문 없이 문제와 선지만 테이블에 배치
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

    # 공통 푸터 마감
    html_content += """\t\t\t</div>
\t\t</div>
\t\t\t</div>
\t</body>
</html>"""

    return html_content


# ==========================================
# [기능 3] Streamlit 웹 대시보드 제어부
# ==========================================

st.set_page_config(page_title="수능형 종합 레이아웃 가공 시스템", layout="wide")

with st.sidebar:
    st.header("⚙️ 통합 설정")
    st.text_input("service_code", value="SVC170")
    st.text_input("track_code", value="RSV_TRK01")
    st.text_input("level_code", value="TO_R_E_SP")

st.title("🧩 수능형 종합 레이아웃 가공 시스템 (태그 고도화 버전)")
st.caption("독해 워드 문서를 업로드하면 19~22번의 colspan='5' 처리 및 23번 이후의 indent 완전 자동화 규칙을 완벽 적용합니다.")

uploaded_file = st.file_uploader("수능 독해형 워드 문서(.docx)를 업로드하세요", type=["docx"])
submit_button = st.button("🚀 문항별 규격 세분화 일괄 ZIP 생성", type="primary")

if uploaded_file is not None and submit_button:
    try:
        with st.spinner("문항 번호를 추적하여 분기형 HTML 태그 구조를 생성하는 중입니다..."):
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
            
        st.success(f"🎉 가공 성공! 총 {len(parsed_data)}문항의 커스텀 레이아웃 맞춤형 ZIP 번들이 완료되었습니다.")
        
        st.subheader("📂 맞춤형 파일 번들 다운로드")
        st.download_button(
            label="📥 세부 가공 완료 통합 ZIP 다운로드",
            data=zip_data,
            file_name="advanced_reading_layout_package.zip",
            mime="application/zip"
        )
        st.info("💡 압축 파일 내부의 19번대 폴더를 열어보시면 수정 요청하신 규격 태그가 완전무결하게 적용된 것을 보실 수 있습니다.")
            
    except Exception as e:
        st.error(f"⚠️ 시스템 패키징 가공 처리 중 예외 발생: {e}")
