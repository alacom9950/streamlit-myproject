import io
import os
import struct
import zipfile
import zlib
from xml.etree import ElementTree as ET

import docx
import olefile
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from pypdf import PdfReader

# .env 파일에서 환경변수 불러오기
load_dotenv()

# 모델명은 여기서 한 번에 바꿀 수 있습니다.
# 참고: 웹검색(web_search) 도구는 Responses API에서 이를 지원하는 모델에서만 동작합니다.
# 만약 웹검색 모드에서 오류가 나면 먼저 이 모델명을 web_search 지원 모델로 바꿔보세요.
MODEL = "gpt-5.4-nano"

# 문서 기반 답변에 사용할 최대 글자 수 (너무 길면 앞부분만 사용)
MAX_DOC_CHARS = 8000

# 한글 문서(HWP/HWPX) 추출 실패 시 안내 문구
HWP_EXTRACT_FAIL_MSG = (
    "해당 한글 문서는 텍스트 추출이 어렵습니다. "
    "PDF 또는 DOCX로 변환 후 다시 업로드해주세요."
)
# 그 외 형식 추출 실패 시 안내 문구
GENERIC_EXTRACT_FAIL_MSG = (
    "문서에서 텍스트를 추출하지 못했습니다. "
    "파일이 손상되지 않았는지 확인하거나 다른 형식으로 변환해 다시 시도해주세요."
)


def _extract_pdf(data):
    reader = PdfReader(io.BytesIO(data))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def _extract_docx(data):
    document = docx.Document(io.BytesIO(data))
    return "\n".join(p.text for p in document.paragraphs)


def _extract_hwpx(data):
    # HWPX는 zip 컨테이너. Contents/section*.xml 안의 텍스트(<...:t>)를 모은다.
    parts = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = sorted(
            n
            for n in zf.namelist()
            if n.startswith("Contents/section") and n.endswith(".xml")
        )
        for name in names:
            root = ET.fromstring(zf.read(name))
            for elem in root.iter():
                tag = elem.tag.rsplit("}", 1)[-1]
                if tag == "p":
                    parts.append("\n")
                elif tag == "t" and elem.text:
                    parts.append(elem.text)
    return "".join(parts)


# HWP 본문에서 8글자(16바이트) 폭을 차지하는 제어 문자 코드들
_HWP_WIDE_CONTROLS = {
    1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15,
    16, 17, 18, 19, 20, 21, 22, 23,
}


def _clean_hwp_paragraph(text):
    # UTF-16 문단 문자열에서 제어 문자를 걸러내고 본문 글자만 남긴다.
    result = []
    i = 0
    length = len(text)
    while i < length:
        code = ord(text[i])
        if code in (10, 13):
            result.append("\n")
            i += 1
        elif code in _HWP_WIDE_CONTROLS:
            i += 8  # 확장/인라인 컨트롤은 8글자를 차지하므로 통째로 건너뜀
        elif code < 32:
            i += 1
        else:
            result.append(text[i])
            i += 1
    return "".join(result)


def _extract_hwp(data):
    # HWP 5.0은 OLE 복합 파일. BodyText/Section* 스트림을 풀어 본문 텍스트를 모은다.
    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        header = ole.openstream("FileHeader").read()
        is_compressed = bool(header[36] & 1)

        section_ids = []
        for entry in ole.listdir():
            if (
                len(entry) == 2
                and entry[0] == "BodyText"
                and entry[1].startswith("Section")
            ):
                section_ids.append(int(entry[1][len("Section"):]))
        section_ids.sort()

        texts = []
        for sid in section_ids:
            raw = ole.openstream(f"BodyText/Section{sid}").read()
            unpacked = zlib.decompress(raw, -15) if is_compressed else raw
            i = 0
            n = len(unpacked)
            while i + 4 <= n:
                rec_header = struct.unpack_from("<I", unpacked, i)[0]
                i += 4
                tag_id = rec_header & 0x3FF
                rec_len = (rec_header >> 20) & 0xFFF
                if rec_len == 0xFFF:  # 길이가 크면 다음 4바이트에 실제 길이가 담김
                    rec_len = struct.unpack_from("<I", unpacked, i)[0]
                    i += 4
                rec_data = unpacked[i:i + rec_len]
                i += rec_len
                if tag_id == 67:  # HWPTAG_PARA_TEXT
                    para = rec_data.decode("utf-16-le", errors="ignore")
                    texts.append(_clean_hwp_paragraph(para))
        return "\n".join(texts)
    finally:
        ole.close()


def extract_text_from_file(uploaded_file):
    # (추출된 텍스트, 오류 안내문) 튜플을 돌려준다. 성공하면 오류 안내문은 빈 문자열.
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    is_hangul = name.endswith((".hwp", ".hwpx"))
    try:
        if name.endswith(".pdf"):
            text = _extract_pdf(data)
        elif name.endswith(".docx"):
            text = _extract_docx(data)
        elif name.endswith(".hwpx"):
            text = _extract_hwpx(data)
        elif name.endswith(".hwp"):
            text = _extract_hwp(data)
        else:
            return "", "지원하지 않는 파일 형식입니다. PDF, DOCX, HWP, HWPX만 업로드할 수 있어요."

        if not text.strip():
            return "", (HWP_EXTRACT_FAIL_MSG if is_hangul else GENERIC_EXTRACT_FAIL_MSG)
        return text, ""
    except Exception:
        return "", (HWP_EXTRACT_FAIL_MSG if is_hangul else GENERIC_EXTRACT_FAIL_MSG)


def build_doc_qa_system(doc_text):
    # 업로드한 문서 내용만 근거로 답하도록 지시하는 시스템 메시지를 만든다.
    context = doc_text[:MAX_DOC_CHARS]
    return (
        "너는 업로드된 문서의 내용만을 근거로 답하는 도우미야. 다음 규칙을 반드시 지켜.\n"
        "- 아래 문서 내용에 있는 정보만으로 한국어로 답한다.\n"
        "- 문서에서 확인할 수 없는 내용은 절대 추측하지 말고, "
        "정확히 '업로드된 문서에서 확인하기 어렵습니다'라고 답한다.\n"
        "- 간결하게 답한다.\n\n"
        "다음은 참고할 문서 내용이야(길면 앞부분 일부만 제공됨):\n"
        "--------\n"
        f"{context}\n"
        "--------"
    )


st.set_page_config(page_title="나의 AI 챗봇", page_icon="🟡", layout="centered")

# 약간의 마감 스타일 (메인 컬러는 .streamlit/config.toml 에서 관리)
st.markdown(
    """
    <style>
      /* 본문 상단 여백을 줄여 깔끔하게 */
      .block-container { padding-top: 2.5rem; max-width: 820px; }
      /* 헤더 영역 */
      .app-header { margin-bottom: 1.2rem; }
      .app-header h1 {
          font-size: 2rem; font-weight: 700; margin: 0;
          letter-spacing: -0.02em;
      }
      .app-header .accent { color: #F5B400; }
      .app-header p { color: #6B7280; margin: 0.25rem 0 0; font-size: 0.95rem; }
      /* 메인 컬러 강조 막대 */
      .app-header .bar {
          width: 48px; height: 4px; border-radius: 999px;
          background: #F5B400; margin-top: 0.7rem;
      }
      /* 채팅 입력창을 살짝 다듬기 */
      [data-testid="stChatInput"] textarea { font-size: 0.98rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# 세련된 헤더
st.markdown(
    """
    <div class="app-header">
      <h1>나의 <span class="accent">AI 챗봇</span></h1>
      <p>무엇이든 물어보세요. 웹검색으로 최신 정보도 찾아드려요.</p>
      <div class="bar"></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# 사이드바: 웹검색 / 문서 기반 답변 옵션
with st.sidebar:
    st.markdown("### ⚙️ 설정")
    use_web_search = st.toggle("웹검색 사용하기")
    st.caption("켜면 최신 웹 정보를 검색해 답변합니다.")

    use_doc_qa = st.toggle("문서 기반 답변 사용하기")
    st.caption("업로드한 문서 내용만 참고해 답변합니다.")
    uploaded_file = st.file_uploader(
        "문서 업로드 (PDF, DOCX, HWP, HWPX)",
        type=["pdf", "docx", "hwp", "hwpx"],
    )

# 업로드된 문서에서 텍스트 추출 (같은 파일은 한 번만 추출하도록 세션에 캐시)
doc_text = ""
doc_error = ""
if uploaded_file is not None:
    file_id = (uploaded_file.name, uploaded_file.size)
    if st.session_state.get("doc_file_id") != file_id:
        with st.spinner("문서에서 텍스트를 추출하고 있어요..."):
            text, err = extract_text_from_file(uploaded_file)
        st.session_state.doc_file_id = file_id
        st.session_state.doc_name = uploaded_file.name
        st.session_state.doc_type = uploaded_file.name.rsplit(".", 1)[-1].upper()
        st.session_state.doc_text = text
        st.session_state.doc_error = err

    doc_text = st.session_state.get("doc_text", "")
    doc_error = st.session_state.get("doc_error", "")

    # 업로드한 문서 정보(파일명/형식/길이)를 사이드바에 표시
    with st.sidebar:
        st.markdown("---")
        st.markdown(f"**파일명**: {st.session_state.doc_name}")
        st.markdown(f"**형식**: {st.session_state.doc_type}")
        st.markdown(f"**추출된 텍스트 길이**: {len(doc_text):,}자")
        if doc_error:
            st.warning(doc_error)
        elif doc_text:
            st.success("문서 준비 완료! 질문해보세요.")

# API Key 불러오기
api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    st.error(
        "OPENAI_API_KEY를 찾을 수 없습니다. "
        "프로젝트 폴더의 .env 파일에 OPENAI_API_KEY를 설정했는지 확인해주세요."
    )
    st.stop()

client = OpenAI(api_key=api_key)

# 웹검색 답변을 간결하고 일정한 형식으로 정리하기 위한 지시문
WEB_SEARCH_INSTRUCTIONS = """\
너는 웹검색 결과를 간결하게 정리해주는 도우미야. 아래 형식을 반드시 지켜서 한국어로 답해.

## 요약
- 핵심 내용을 5줄 이내로 먼저 정리한다.

## 조사 결과
- 주요 내용을 bullet point로 정리한다.
- 가능하면 각 항목 끝에 참고 링크를 [제목](URL) 형태로 함께 붙인다.

## 내 프로젝트에 적용할 수 있는 점
- 실제로 적용해볼 만한 점을 딱 2개만 제안한다.

전체적으로 장황하지 않게, 꼭 필요한 내용만 담아 간결하게 작성해."""


# 웹검색 응답에서 참고한 출처(URL) 목록을 뽑아냅니다.
def extract_citations(response):
    citations = []
    seen = set()
    for item in getattr(response, "output", []) or []:
        if getattr(item, "type", None) != "message":
            continue
        for content in getattr(item, "content", []) or []:
            for ann in getattr(content, "annotations", []) or []:
                if getattr(ann, "type", None) != "url_citation":
                    continue
                url = getattr(ann, "url", None)
                if not url or url in seen:
                    continue
                seen.add(url)
                citations.append((getattr(ann, "title", "") or url, url))
    return citations

# 대화 기록 초기화 (이전 대화가 화면에 계속 남아 있도록 세션에 저장)
if "messages" not in st.session_state:
    st.session_state.messages = []

# 지금까지의 대화 내용 화면에 표시
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# 사용자 입력 받기
if prompt := st.chat_input("무엇이든 물어보세요!"):
    # 사용자 질문 저장 및 표시
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # AI 답변 생성
    with st.chat_message("assistant"):
        try:
            if use_doc_qa:
                # 문서 기반 답변 모드 (웹검색보다 우선 적용)
                if not doc_text:
                    # 문서 미업로드(또는 추출 실패) 시 안내만 표시
                    answer = doc_error or "먼저 문서를 업로드해주세요."
                    st.markdown(answer)
                else:
                    with st.spinner("문서를 살펴보고 있어요..."):
                        system_message = {
                            "role": "system",
                            "content": build_doc_qa_system(doc_text),
                        }
                        stream = client.chat.completions.create(
                            model=MODEL,
                            messages=[system_message] + st.session_state.messages,
                            stream=True,
                        )

                        def token_stream():
                            for chunk in stream:
                                delta = chunk.choices[0].delta.content
                                if delta:
                                    yield delta

                        answer = st.write_stream(token_stream())
            elif use_web_search:
                # 웹검색 모드: Responses API의 web_search 도구로 답변
                with st.spinner("웹을 검색하고 있어요..."):
                    response = client.responses.create(
                        model=MODEL,
                        tools=[{"type": "web_search"}],
                        instructions=WEB_SEARCH_INSTRUCTIONS,
                        input=st.session_state.messages,
                    )

                answer = response.output_text

                # 참고한 출처가 있으면 답변 아래에 링크로 덧붙여 함께 저장/표시
                citations = extract_citations(response)
                if citations:
                    sources = "\n".join(
                        f"- [{title}]({url})" for title, url in citations
                    )
                    answer = f"{answer}\n\n**참고한 출처**\n{sources}"

                st.markdown(answer)
            else:
                # 기존 챗봇 모드: 스트리밍으로 답변
                with st.spinner("답변을 생각하고 있어요..."):
                    stream = client.chat.completions.create(
                        model=MODEL,
                        messages=st.session_state.messages,
                        stream=True,
                    )

                    # 스트리밍으로 받은 텍스트 조각을 한 글자씩 화면에 타이핑하듯 표시
                    def token_stream():
                        for chunk in stream:
                            delta = chunk.choices[0].delta.content
                            if delta:
                                yield delta

                    answer = st.write_stream(token_stream())

            # AI 답변도 대화 기록에 저장
            st.session_state.messages.append(
                {"role": "assistant", "content": answer}
            )
        except Exception as e:
            st.error(
                "죄송합니다. 답변을 생성하는 중에 문제가 발생했습니다.\n\n"
                "잠시 후 다시 시도하시거나, API Key와 인터넷 연결을 확인해주세요."
            )
            # 개발자가 원인을 확인할 수 있도록 상세 오류는 접어서 표시
            with st.expander("자세한 오류 내용 보기"):
                st.write(str(e))