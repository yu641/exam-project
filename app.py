import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime
import numpy as np

# --- 사용자 정의 함수 임포트 ---
# 제공해주신 .py 파일들이 app.py와 동일한 폴더에 있다고 가정합니다.
try:
    from exam_functions import (
        grade_exam, 
        analyze_weakness_from_graded_file,
        generate_exam_7_passages_from_db, # 시험지 생성(랜덤)
        generate_exam_irt_weakness        # 시험지 생성(IRT)
    )
    from run_CREATE_DASHBOARD import create_dashboard
except ImportError:
    st.error("오류: `exam_functions.py` 또는 `run_CREATE_DASHBOARD.py` 파일을 찾을 수 없습니다. `app.py`와 동일한 폴더에 있는지 확인하세요.")
    st.stop()

# --- 상수 및 디렉토리 설정 ---
TEMP_DIR = "./temp"
# OUTPUT_DIR은 사이드바에서 설정한 값을 사용합니다.
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs("./answers", exist_ok=True) # 답안 예시 폴더

# --- 헬퍼 함수 ---
def save_uploaded_file(uploaded_file, directory=TEMP_DIR):
    """업로드된 파일을 임시 디렉토리에 저장하고 경로를 반환합니다."""
    if uploaded_file is not None:
        path = os.path.join(directory, uploaded_file.name)
        with open(path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return path
    return None

def read_file_for_download(file_path):
    """다운로드 버튼을 위해 파일을 읽습니다."""
    try:
        with open(file_path, "rb") as f:
            return f.read()
    except FileNotFoundError:
        st.error(f"파일을 찾을 수 없습니다: {file_path}")
        return None

# --- Streamlit UI ---

st.set_page_config(layout="wide")
st.sidebar.title("시험 분석 시스템")

# --- 1. 경로 설정 (사이드바) ---
st.sidebar.header("폴더 및 파일 경로 설정")
st.sidebar.info("앱이 실행되는 위치 기준의 상대 경로를 사용하세요.")

BASE_DIR = st.sidebar.text_input(
    "지문/문제 폴더 경로 (BASE_DIR)", 
    "data",
    help="'지문' 폴더와 '문제' 폴더가 들어있는 상위 폴더입니다. (예: ./data)"
)

# DB 파일 경로
DB_PATH = "data/db_with_irt_from_distractors.xlsx"
st.sidebar.success(f"DB 파일: {DB_PATH}") # 사용자에게 고정된 경로를 알려줍니다.

OUTPUT_DIR = st.sidebar.text_input(
    "출력 폴더 (OUTPUT_DIR)", 
    "output",
    help="생성된 시험지, 메타파일, 채점 결과, 취약점 파일이 저장될 폴더입니다."
)

# 앱 실행 시 출력 폴더 생성
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- 2. 페이지 선택 (사이드바) ---
page = st.sidebar.radio("메뉴", ["시험지 생성", "채점 및 취약점 분석", "대시보드 생성"])

st.sidebar.header("사용 안내")
st.sidebar.warning(
    "**시험지 생성 (로컬 전용)**\n"
    "'시험지 생성' 메뉴는 **MS Word가 설치된 Windows PC**에서 로컬로 Streamlit을 실행할 때만 동작합니다.\n\n",
    icon="💻"
)
st.sidebar.info(
    "**1. 시험지 생성**\n"
    "'시험지 생성' 메뉴에서 모드와 학생 정보를 입력하고 시험지를 생성합니다.",
    icon="📝"
)
st.sidebar.info(
    "**2. 채점**\n"
    "'채점' 메뉴에서 '메타파일'과 '학생 답안'을 업로드하여 채점합니다.",
    icon="📄"
)
st.sidebar.info(
    "**3. 분석**\n"
    "'대시보드' 메뉴에서 '채점 완료 파일'을 업로드하여 성적을 분석합니다.",
    icon="📊"
)

# ==============================================================================
# 페이지 1: 시험지 생성
# ==============================================================================
if page == "시험지 생성":
    st.header("1. 시험지 생성 (로컬 Windows 전용)")
    st.warning(
        "이 기능은 **MS Word가 설치된 Windows PC**에서 로컬로 실행할 때만 정상 동작합니다. "
        "웹 서버(Streamlit Cloud 등)에서는 Word 파일(.docx)을 생성할 수 없습니다.",
        icon="⚠️"
    )
    
    st.subheader("학생 정보 입력")
    col1, col2 = st.columns(2)
    with col1:
        student_id = st.text_input("학생 ID", "S001")
    with col2:
        student_name = st.text_input("학생 이름", "김철수")

    st.subheader("시험 모드 선택")
    mode = st.radio("생성할 시험지 모드를 선택하세요.", ["RANDOM (첫 사용자용)", "IRT (맞춤형)"], horizontal=True)

    user_theta = 0.0
    if mode == "IRT (맞춤형)":
        user_theta = st.number_input("학생 능력치 (Theta)", min_value=-3.0, max_value=3.0, value=0.3, step=0.1)
        st.info(f"IRT 모드 선택됨: {student_id} 학생의 Theta 값 {user_theta}를 사용합니다.\n"
                f"취약점 파일: `{os.path.join(OUTPUT_DIR, f'user_weakness_{student_id}.xlsx')}` 를 참조합니다.")

    if st.button("시험지 생성 시작하기", type="primary"):
        
        # --- 경로 검증 ---
        if not os.path.exists(DB_PATH):
            st.error(f"DB 파일을 찾을 수 없습니다. (경로: {DB_PATH})")
            st.stop()
        if not os.path.exists(BASE_DIR):
            st.error(f"지문/문제 폴더를 찾을 수 없습니다. (경로: {BASE_DIR})")
            st.stop()
        if not os.path.exists(os.path.join(BASE_DIR, "지문")) or not os.path.exists(os.path.join(BASE_DIR, "문제")):
            st.warning(f"'{BASE_DIR}' 폴더 내에 '지문' 또는 '문제' 폴더가 있는지 확인하세요.")

        
        with st.spinner(f"{mode} 모드로 시험지 생성 중... (MS Word가 실행될 수 있습니다)"):
            gen_result = None
            try:
                if mode == "RANDOM (첫 사용자용)":
                    # generate_exam_7_passages_from_db 함수는 num_passages, num_problems_per_passage 인자를 받지 않으므로 제거합니다.
                    gen_result = generate_exam_7_passages_from_db(
                        db_path=DB_PATH,
                        base_dir=BASE_DIR,
                        title="[첫 사용자용] 국어 영역 시험지",
                        subtitle=f"{student_name} 학생",
                        output_dir=OUTPUT_DIR,
                        student_id=student_id,
                        student_name=student_name,
                        # num_passages=7,                 # <-- 이 인자가 오류의 원인입니다. (제거)
                        # num_problems_per_passage=4,     # <-- 이 인자도 제거합니다.
                        two_columns=True
                    )
                
                elif mode == "IRT (맞춤형)":
                    weakness_file_path = os.path.join(OUTPUT_DIR, f"user_weakness_{student_id}.xlsx")
                    if not os.path.exists(weakness_file_path):
                        st.warning(f"취약점 파일({weakness_file_path})을 찾을 수 없습니다. IRT 모드이지만 취약점 가중치 없이 생성됩니다.")
                    
                    # [참고] generate_exam_irt_weakness 함수는 해당 인자를 받으므로 그대로 둡니다.
                    gen_result = generate_exam_irt_weakness(
                        db_path=DB_PATH,
                        base_dir=BASE_DIR,
                        user_weakness_path=weakness_file_path,
                        user_theta=user_theta,
                        title="[맞춤형] 국어 영역 시험지",
                        subtitle=f"{student_name}님 취약점 보완 (Theta={user_theta})",
                        num_passages=7,
                        num_problems_per_passage=4,
                        weak_passage_target_prop=0.6,
                        weak_problem_boost=1.5,
                        two_columns=True,
                        output_dir=OUTPUT_DIR,
                        student_id=student_id,
                        student_name=student_name
                    )

                # --- 결과 처리 ---
                if gen_result and isinstance(gen_result, tuple) and len(gen_result) >= 3:
                    # 튜플의 순서가 (doc_path, meta_path, exam_id)라고 가정합니다.
                    doc_path, meta_path, exam_id = gen_result[0], gen_result[1], gen_result[2]
                    
                    # doc_path 또는 meta_path가 None인지 확인 (TypeError 방지)
                    if doc_path is None or meta_path is None:
                        st.error("시험지 생성에 실패했습니다 (파일 경로가 반환되지 않았습니다).")
                        st.error("MS Word가 정상적으로 실행되었는지, 권한 문제가 없는지, 백그라운드에서 실행 중인지 확인하세요.")
                        st.info(f"반환된 값: doc_path={doc_path}, meta_path={meta_path}")
                        # 메타 파일이라도 생성되었으면 다운로드 링크 제공
                        if meta_path:
                            meta_filename = os.path.basename(meta_path)
                            st.download_button(
                                label=f"메타파일 (.xlsx) (생성됨)\n({meta_filename})",
                                data=read_file_for_download(meta_path),
                                file_name=meta_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        # 정상 처리
                        st.success(f"시험지 생성 완료! (시험 ID: {exam_id})")
                        st.info(f"`{OUTPUT_DIR}` 폴더에 파일이 저장되었습니다.")

                        st.subheader("생성된 파일 다운로드")
                        doc_filename = os.path.basename(doc_path)
                        meta_filename = os.path.basename(meta_path)

                        dl_col1, dl_col2 = st.columns(2)
                        with dl_col1:
                            st.download_button(
                                label=f"1. 시험지 (.docx)\n({doc_filename})",
                                data=read_file_for_download(doc_path),
                                file_name=doc_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        with dl_col2:
                            st.download_button(
                                label=f"2. 메타파일 (.xlsx)\n({meta_filename})",
                                data=read_file_for_download(meta_path),
                                file_name=meta_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    
                else:
                    st.error("시험지 생성에 실패했습니다. 터미널(콘솔)의 로그를 확인하세요.")
                    st.error("오류의 원인이 'win32com' 또는 'Word' 관련이라면, MS Word가 설치되어 있는지, Windows 환경이 맞는지 확인하세요.")

            except Exception as e:
                st.error(f"시험지 생성 중 오류 발생: {e}")
                st.exception(e)
                if "win32com" in str(e) or "pywintypes" in str(e):
                    st.error("오류 상세: 'win32com' 라이브러리 관련 문제입니다. MS Word가 설치된 Windows 환경에서만 이 기능을 사용할 수 있습니다.")

# ==============================================================================
# 페이지 2: 채점 및 취약점 분석
# ==============================================================================
elif page == "채점 및 취약점 분석":
    st.header("2. 채점 및 취약점 분석")
    st.info("'시험지 생성' 단계에서 만들어진 '시험 메타파일'과 학생이 작성한 '답안 파일'을 업로드하세요.")

    col1, col2 = st.columns(2)
    with col1:
        exam_meta_file = st.file_uploader("1. 시험 메타파일 (.xlsx)", type="xlsx", help="`시험지 생성` 시 output 폴더에 생성된 `S001_..._1.xlsx`과 같은 파일")
    
    with col2:
        answer_sheet_file = st.file_uploader("2. 학생 답안 파일 (.xlsx)", type="xlsx", help="학생이 답을 입력한 엑셀 파일. 첫 번째 열에 답이 있어야 합니다.")

    if st.button("채점 시작하기", type="primary", disabled=(not exam_meta_file or not answer_sheet_file)):
        
        # 1. 업로드된 파일 임시 저장
        temp_meta_path = save_uploaded_file(exam_meta_file)
        temp_answers_path = save_uploaded_file(answer_sheet_file)

        if temp_meta_path and temp_answers_path:
            with st.spinner("채점 및 취약점 분석 중..."):
                try:
                    # 2. 채점 실행
                    grade_result = grade_exam(
                        exam_xlsx_path=temp_meta_path,
                        interactive=False, # 파일 업로드 방식 사용
                        answers_xlsx_path=temp_answers_path,
                        output_dir=OUTPUT_DIR # 통합된 출력 폴더 사용
                    )
                    
                    if not grade_result or not grade_result.get("graded_path"):
                        st.error("채점에 실패했습니다. 터미널 로그를 확인하세요.")
                        st.stop()

                    st.success(f"채점 완료! **{grade_result['score']}점** ({grade_result['correct']} / {grade_result['total']})")

                    # 3. 취약점 분석 실행
                    updated_weakness_file = analyze_weakness_from_graded_file(
                        graded_xlsx_path=grade_result["graded_path"], # 채점 결과 파일
                        output_dir=OUTPUT_DIR, # 통합된 출력 폴더 사용
                        passage_threshold=70.0,
                        problem_threshold=60.0
                    )

                    if updated_weakness_file:
                        st.success(f"취약점 분석 및 갱신 완료!")
                    else:
                        st.warning("취약점 분석에 실패했습니다.")
                        st.stop()

                    # 4. 결과 파일 다운로드 버튼 제공
                    st.subheader("결과 파일 다운로드")
                    
                    # 파일 경로에서 파일 이름만 추출
                    graded_filename = os.path.basename(grade_result["graded_path"])
                    result_filename = os.path.basename(grade_result["result_path"])
                    weakness_filename = os.path.basename(updated_weakness_file)

                    dl_col1, dl_col2, dl_col3 = st.columns(3)
                    with dl_col1:
                        st.download_button(
                            label=f"1. 상세 채점 파일\n({graded_filename})",
                            data=read_file_for_download(grade_result["graded_path"]),
                            file_name=graded_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with dl_col2:
                        st.download_button(
                            label=f"2. 요약 결과 파일\n({result_filename})",
                            data=read_file_for_download(grade_result["result_path"]),
                            file_name=result_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with dl_col3:
                        st.download_button(
                            label=f"3. 갱신된 취약점 파일\n({weakness_filename})",
                            data=read_file_for_download(updated_weakness_file),
                            file_name=weakness_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                except Exception as e:
                    st.error(f"실행 중 오류가 발생했습니다: {e}")
                    st.exception(e)
                finally:
                    # 임시 파일 삭제
                    if os.path.exists(temp_meta_path): os.remove(temp_meta_path)
                    if os.path.exists(temp_answers_path): os.remove(temp_answers_path)

# ==============================================================================
# 페이지 3: 대시보드 생성
# ==============================================================================
elif page == "대시보드 생성":
    st.header("3. 대시보드 생성")
    st.info("'채점' 단계에서 생성된 상세 채점 파일(`..._graded.xlsx`)을 업로드하세요.")
    
    graded_file = st.file_uploader("상세 채점 파일 (..._graded.xlsx)", type="xlsx", help="`채점` 메뉴에서 다운로드한 '1. 상세 채점 파일'입니다.")

    if st.button("대시보드 생성 및 보기", type="primary", disabled=(not graded_file)):
        
        # 0. DB 파일 존재 여부 확인
        if not os.path.exists(DB_PATH):
            st.error(f"DB 파일을 찾을 수 없습니다. (경로: {DB_PATH})")
            st.stop()

        # 1. 임시 파일 저장
        temp_graded_path = save_uploaded_file(graded_file)
        dashboard_output_path = os.path.join(TEMP_DIR, f"DASHBOARD_{graded_file.name}")

        if temp_graded_path:
            with st.spinner("대시보드 생성 중..."):
                try:
                    # 2. 대시보드 생성 (Excel 파일)
                    create_dashboard(
                        graded_path=temp_graded_path,
                        db_path=DB_PATH,
                        output_path=dashboard_output_path
                    )
                    st.success("대시보드 엑셀 파일 생성 완료!")

                    # 3. 생성된 대시보드 엑셀 파일 다운로드 버튼
                    st.download_button(
                        label="대시보드 Excel 파일 다운로드",
                        data=read_file_for_download(dashboard_output_path),
                        file_name=os.path.basename(dashboard_output_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                finally:
                    # 임시 파일 삭제
                    if os.path.exists(temp_graded_path): os.remove(temp_graded_path)
                    if os.path.exists(dashboard_output_path): os.remove(dashboard_output_path)
