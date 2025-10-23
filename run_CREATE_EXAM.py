import os
from exam_functions import generate_exam_7_passages_from_db, generate_exam_irt_weakness

# --- 설정 (필수) ---
# (중요) "RANDOM": 첫 사용자용, "IRT": 맞춤형
MODE = "IRT" 

BASE_DIR = r".\data"           # 지문/문제 docx 루트
DB_PATH  = r".\data\db_with_irt_from_distractors.xlsx"   # DB (IRT 모드 시 irt_... 컬럼 필수)
WEAKNESS_DB_DIR = r".\output"    # 취약점 파일(user_weakness...) 저장/로드 폴더
OUTPUT_DIR = r".\output"       # 시험지(docx), 메타(xlsx) 저장 폴더

STUDENT_ID = "S001"            # (필수) 응시할 학생 ID
STUDENT_NAME = "김철수"        # (필수) 응시할 학생 이름

# --- IRT 모드일 때만 필요한 설정 ---
USER_THETA = 0.3               # (IRT 필수) 학생의 현재 능력치
# ------------------------------

if __name__ == "__main__":

    if MODE == "RANDOM":
        print(f"--- {STUDENT_NAME}({STUDENT_ID})님 첫 사용자용 랜덤 시험 생성 ---")
        
        gen_result = generate_exam_7_passages_from_db(
            db_path=DB_PATH,
            base_dir=BASE_DIR,
            title="[진단] 국어 영역 시험지",
            subtitle=f"{STUDENT_NAME}님 진단 평가",
            two_columns=True,
            output_dir=OUTPUT_DIR,
            student_id=STUDENT_ID,
            student_name=STUDENT_NAME
        )
    
    elif MODE == "IRT":
        print(f"--- {STUDENT_NAME}({STUDENT_ID})님 맞춤형 시험 생성 (Theta={USER_THETA}) ---")
        
        # 학생 ID를 기반으로 취약점 파일 경로 자동 설정
        WEAKNESS_FILE_PATH = os.path.join(WEAKNESS_DB_DIR, f"user_weakness_{STUDENT_ID}.xlsx")
        print(f"사용할 취약점 파일: {WEAKNESS_FILE_PATH}")

        gen_result = generate_exam_irt_weakness(
            db_path=DB_PATH,
            base_dir=BASE_DIR,
            user_weakness_path=WEAKNESS_FILE_PATH,
            user_theta=USER_THETA,
            title="[맞춤형] 국어 영역 시험지",
            subtitle=f"{STUDENT_NAME}님 취약점 보완 (Theta={USER_THETA})",
            num_passages=7,
            num_problems_per_passage=4,
            weak_passage_target_prop=0.6,
            weak_problem_boost=1.5,
            two_columns=True,
            output_dir=OUTPUT_DIR,
            
            student_id=STUDENT_ID,
            student_name=STUDENT_NAME
        )
        
    else:
        print(f"오류: MODE 변수('{MODE}')가 잘못되었습니다. 'RANDOM' 또는 'IRT'로 설정하세요.")
        gen_result = None

    # --- 결과 출력 ---
    if gen_result and gen_result[1]:
        exam_word_path, exam_meta_path, exam_id = gen_result
        print("\n--- 시험지 생성 성공 ---")
        print(f"시험 ID: {exam_id}")
        print(f"시험지(Word): {exam_word_path}")
        print(f"메타(Excel): {exam_meta_path}")
        print("\n학생이 시험을 푼 후, 답안 파일과 메타 파일을 'run_GRADE_EXAM.py'로 채점하세요.")
    else:
        print("\n--- 시험지 생성 실패 ---")
        