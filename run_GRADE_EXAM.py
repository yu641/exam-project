import os
from exam_functions import grade_exam, analyze_weakness_from_graded_file

# --- 설정 (필수) ---

# (필수) 채점할 시험지의 메타 파일 (시험지 생성 시 output 폴더에 생성된 ...exam_id.xlsx 파일)
EXAM_FILE_TO_GRADE = r".\output\EX20251021T072918-S001-EXAM.xlsx" 

# (필수) 학생이 제출한 답안지 (엑셀 파일 첫 열에 1,2,3,4,5 입력)
# None으로 설정 시, interactive=True (직접 입력) 모드로 실행됩니다.
ANSWERS_FILE_PATH = r".\answers\S001_EX20251021T072918-S001-EXAM_answers.xlsx" 

# (필수) 채점 결과(..._graded.xlsx)를 저장할 폴더
GRADE_OUTPUT_DIR = r".\output"

# (필수) 분석된 취약점 파일(user_weakness_...xlsx)을 저장할 폴더
WEAKNESS_DB_DIR = r".\output"
# --------------------

if __name__ == "__main__":
    
    if not os.path.exists(EXAM_FILE_TO_GRADE):
        print(f"오류: 채점할 메타 파일을 찾을 수 없습니다. ({EXAM_FILE_TO_GRADE})")
    else:
        print(f"--- 채점 및 분석 시작 ---")
        print(f"대상 메타 파일: {EXAM_FILE_TO_GRADE}")
        
        # 1. 채점
        # ANSWERS_FILE_PATH가 None이면 interactive=True가 됨
        grade_result = grade_exam(
            exam_xlsx_path=EXAM_FILE_TO_GRADE,
            interactive=(ANSWERS_FILE_PATH is None),
            answers_xlsx_path=ANSWERS_FILE_PATH,
            output_dir=GRADE_OUTPUT_DIR
        )
        
        # 2. 분석 (채점 성공 시)
        if grade_result and grade_result.get("graded_path"):
            
            print(f"\n--- 취약점 분석 및 갱신 시작 ---")
            
            updated_weakness_file = analyze_weakness_from_graded_file(
                graded_xlsx_path=grade_result["graded_path"], # 채점 결과 파일
                output_dir=WEAKNESS_DB_DIR, # 취약점 파일 저장 위치
                passage_threshold=70.0,
                problem_threshold=60.0
            )
            
            if updated_weakness_file:
                print(f"\n{grade_result['student_id']} 학생의 취약점 파일 갱신 완료:")
                print(f"{updated_weakness_file}")
            else:
                print("\n--- 취약점 파일 갱신 실패 ---")
        else:
            print("\n--- 채점 실패 (분석 건너뜀) ---")