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
    
    # 1차 순회: 최상위 body 기준 문단과 표(지문 박스)를 순서대로 수집 (이중 파싱 방지)
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
                
        elif element.tag.endswith('tbl'):  # 📦 표 (19번 이후의 박스형 지문)
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

    # 2차 순회: 수능 독해 레이아웃 규칙에 맞춰 그룹화
    for item in all_elements:
        # [A] 박스형 지문(표)을 만났을 때 처리
        if item["type"] == "table":
            if current_q is not None:
                for t_line in item["lines"]:
                    # 밑줄 공백 표현 변환
                    converted_t_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', t_line)
                    current_q["passage_lines"].append(converted_t_line)
            continue

        line = item["text"].strip()

        # [B] 새로운 문제 번호 등장 시 마감 및 새 방 개설
        q_match = q_pattern.match(line)
        if q_match:
            if current_q is not None:
                questions.append(current_q)
            current_q = {
                "num": q_match.group(1),
                "title": q_match.group(2),
                "passage_lines": [],  # <div class="passage"> 내부의 <p class="dialog">로 들어갈 지문들
                "options": []
            }
            continue

        if current_q is None:
            continue

        # 교사용 정답 및 해설 라인 패스
        if line.startswith("정답:") or "→" in line:
            continue
            
        # 대괄호로 묶인 문제 번호 안내선 패스 (예: [26~28])
        if line.startswith("[") and "]" in line:
            continue

        # [C] 선택지 보기(①~⑤)를 만났을 때
        opt_match = opt_pattern.match(line)
        if opt_match:
            label_char = opt_match.group(1)
            opt_text = opt_match.group(2)
            
            label_map = {"①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5"}
            label_num = label_map.get(label_char, "1")
            
            # 보기 내부 연속 공백 정렬 복원
            processed_opt = re.sub(r'\s{2,}', ' &nbsp;&nbsp;&nbsp;&nbsp; ', opt_text)
            current_q["options"].append({
                "char": label_char,
                "num": label_num,
                "text": processed_opt
            })
            continue

        # [D] 선지 전이나 사이에 있는 모든 지문/단어 주석은 p class="dialog"로 흡수
        converted_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', line)
        current_q["passage_lines"].append(converted_line)

    if current_q is not None:
        questions.append(current_q)

    return questions


# ==========================================
# [기능 2] 수능 독해 규격 맞춤형 단일 HTML 생성
# ==========================================
def generate_reading_html(q):
    # 제공해주신 수능 독해 표준 스킨 템플릿 선언
    html_content = """<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta charset="UTF-8" name="viewport" content="width=device-width, target-densitydpi=device-dpi" />
<meta HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no,minimum-scale=1.0,maximum-scale=1.0,target-densitydpi=device-dpi">
<title>Online Achievement Test</title>
<link rel="stylesheet" href="../../../common/css/common.css">
<link rel="stylesheet" href="../../../common/css/style.css">
<link rel="stylesheet" href="../../../common_ex/css/style.css">
<link rel="stylesheet" href="../../../common_ex/css/scroll.css">
<script type="text/javascript" charset="UTF-8" src="../../../common/js/jquery.js"></script>
<script type="text/javascript" charset="UTF-8" src="../../../common/js/common.js"></script>
<script type="text/javascript" charset="UTF-8" src="../../../common_ex/js/common.js"></script>
</head>
<body>
<div class="pageWrap">
\t<div class="listening_desc_box"><b>Reading Comprehension</b> 질문을 읽고 물음에 답하시오.</div>
\t<div id="R_question" class="STSection">
\t\t<div class="pageConts">
"""

    # 지시문 내 밑줄 태그가 깨지지 않도록 안전 변환 및 복원
    safe_title = html.escape(q['title'])
    safe_title = safe_title.replace("&lt;u&gt;", "<u>").replace("&lt;/u&gt;", "</u>")

    # 상단 문제 및 지시문 배치
    html_content += f"""\t\t\t<div class="q_box">
\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer="3" gravity="top|left">
\t\t\t\t\t<tr>
\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t<td>{safe_title} <br></td>
\t\t\t\t\t</tr>\n"""
    
    # 💡 [핵심 요청 구현] 지문 영역을 <div class="passage">로 감싸고, 각 문단을 <p class="dialog">로 처리
    if q["passage_lines"]:
        html_content += """\t\t\t\t\t<tr>
\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t<td>
\t\t\t\t\t\t\t<div class="passage">\n"""
        
        for p_line in q["passage_lines"]:
            html_content += f"\t\t\t\t\t\t\t\t<p class=\"dialog\">{p_line}</p>\n"
            
        html_content += """\t\t\t\t\t\t\t</div>
\t\t\t\t\t\t</td>
\t\t\t\t\t</tr>\n"""
        
    # 하단 선택지(①~⑤) 테이블 형태로 정렬 배치
    for opt in q["options"]:
        html_content += f"""\t\t\t\t\t<tr>
\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t<td><span class="STChoice" remarkable="true" label="{opt['num']}"><span class="label">{opt['char']}</span> {opt['text']} </span></td>
\t\t\t\t\t</tr>\n"""
        
    html_content += """\t\t</table>
\t\t\t</div>\n"""

    html_content += """\t\t</div>
\t</div>
</div>
</body>
</html>"""

    return html_content


# ==========================================
# [기능 3] Streamlit 웹 화면 레이아웃 및 다운로드 제어
# ==========================================

st.set_page_config(page_title="수능 독해형 자동 HTML 추출 시스템", layout="wide")

with st.sidebar:
    st.header("⚙️ 독해 고정값 설정")
    st.text_input("service_code", value="SVC170")
    st.text_input("track_code", value="RSV_TRK01")
    st.text_input("top_cors_id", value="1879")
    st.text_input("level_code", value="TO_R_E_SP")
    st.text_input("component_code", value="COM170")
    st.text_input("book_code", value="SVC170")
    st.text_input("act_name", value="Reading")

st.title("📚 수능형 독해 자동 HTML 추출 시스템")
st.caption("수능 독해 정기평가 워드 문서를 업로드하면 일반 지문 및 박스형 지문을 정밀하게 감지하여 규격화된 HTML 패키지로 분할 생성합니다.")

uploaded_file = st.file_uploader("독해 워드 파일(.docx)을 업로드하세요", type=["docx"])
submit_button = st.button("🚀 독해 문항 폴더 구조로 분할 변환하기", type="primary")

if uploaded_file is not None and submit_button:
    try:
        with st.spinner("수능형 지문 구조 파악 및 선지 전 dialog 변환 작업을 진행 중입니다..."):
            parsed_data = parse_reading_docx(uploaded_file)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for q in parsed_data:
                    q_num = q["num"]
                    if not q_num:
                        continue
                    
                    single_html = generate_reading_html(q)
                    folder_file_path = f"{q_num}/test.html"
                    zip_file.writestr(folder_file_path, single_html)
            
            zip_data = zip_buffer.getvalue()
            
        st.success(f"🎉 변환 완료! 총 {len(parsed_data)}개의 수능 독해 문항 지문과 선지를 성공적으로 연동했습니다.")
        
        st.subheader("📂 독해 패키지 다운로드")
        st.download_button(
            label="📥 독해 번호별 폴더 압축파일(.zip) 다운로드",
            data=zip_data,
            file_name="reading_questions_folders.zip",
            mime="application/zip"
        )
        st.info(f"💡 일반 지문 및 19번 이후 박스 지문 모두 <div class='passage'> 하위의 <p class='dialog'> 구조로 완벽 가공되었습니다.")
            
    except Exception as e:
        st.error(f"⚠️ 독해 변환 시스템 처리 중 예외 발생: {e}")
