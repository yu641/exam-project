import pandas as pd
import numpy as np
import os
import glob

# --- 설정 ---
# (필수) 채점 완료된 파일(..._graded.xlsx)들이 모여있는 폴더
GRADED_RESULTS_DIR = r".\output_irt" 

# (필수) IRT 값을 갱신할 마스터 DB 파일. 이 파일에 덮어씁니다.
MASTER_DB_PATH = r".\db_with_irt_from_distractors.xlsx"

# (필수) 백업 파일 저장 경로
BACKUP_DB_PATH = r".\db_backup.xlsx"
# ------------

print("--- IRT 값 갱신 스크립트 시작 ---")

# 1. 원본 DB 백업
if os.path.exists(MASTER_DB_PATH):
    print(f"원본 마스터 DB 파일을 '{BACKUP_DB_PATH}'(으)로 백업합니다.")
    master_df_for_backup = pd.read_excel(MASTER_DB_PATH)
    master_df_for_backup.to_excel(BACKUP_DB_PATH, index=False)
else:
    print(f"오류: 마스터 DB 파일 '{MASTER_DB_PATH}'을(를) 찾을 수 없습니다.")
    exit()

# 2. 모든 채점 결과 파일에서 답안 데이터 집계
graded_files = glob.glob(os.path.join(GRADED_RESULTS_DIR, "*_graded.xlsx"))

if not graded_files:
    print(f"오류: '{GRADED_RESULTS_DIR}' 폴더에서 채점된 파일을 찾을 수 없습니다.")
    exit()

print(f"총 {len(graded_files)}개의 채점 파일에서 데이터를 집계합니다.")

all_answers = []
for file in graded_files:
    try:
        # 'answers' 시트에서 문제 ID와 학생 답안을 읽어옴
        df_ans = pd.read_excel(file, sheet_name='answers')
        all_answers.append(df_ans[['problem_id', 'student_answer_num']])
    except Exception as e:
        print(f"경고: '{file}' 파일 처리 중 오류 발생 (건너뜀): {e}")

if not all_answers:
    print("오류: 유효한 답안 데이터를 집계하지 못했습니다.")
    exit()

# 모든 답안 데이터를 하나의 데이터프레임으로 합침
df_aggregated = pd.concat(all_answers, ignore_index=True)

# 'student_answer_num'이 숫자가 아닌 경우를 대비해 변환
df_aggregated['student_answer_num'] = pd.to_numeric(df_aggregated['student_answer_num'], errors='coerce')
df_aggregated.dropna(subset=['problem_id', 'student_answer_num'], inplace=True)
df_aggregated['student_answer_num'] = df_aggregated['student_answer_num'].astype(int)

print(f"총 {len(df_aggregated)}개의 유효 응답을 집계했습니다.")

# 3. 문제 ID별, 선지별 응답 횟수 계산
# crosstab을 사용하여 각 문제(index)별로 각 선지(columns)를 몇 번 선택했는지 카운트
response_counts = pd.crosstab(df_aggregated['problem_id'], df_aggregated['student_answer_num'])

# 1~5번 선지 컬럼이 없는 경우를 대비해 0으로 채워서 추가
for i in range(1, 6):
    if i not in response_counts.columns:
        response_counts[i] = 0
response_counts = response_counts[[1, 2, 3, 4, 5]] # 순서 고정

# 문제별 총 응답 횟수
total_responses = response_counts.sum(axis=1)

# 선지별 선택률 계산
response_rates = response_counts.div(total_responses, axis=0)
response_rates.columns = [f'선지정답률_{i}' for i in range(1, 6)]
response_rates.reset_index(inplace=True)
response_rates.rename(columns={'problem_id':'문제id'}, inplace=True)

print(f"{len(response_rates)}개 문항에 대한 새로운 선지별 정답률을 계산했습니다.")

# 4. 새로운 IRT 값 계산 (기존 로직 재사용)
master_df = pd.read_excel(MASTER_DB_PATH)
new_irt_values = []

# response_rates에 있는 문제들에 대해서만 IRT 값 계산
for index, row in response_rates.iterrows():
    problem_id = row['문제id']
    
    # 마스터 DB에서 해당 문제의 정답 정보 가져오기
    correct_answer_info = master_df.loc[master_df['문제id'] == problem_id, '정답']
    if correct_answer_info.empty:
        continue # 마스터 DB에 없는 문제면 건너뜀
    
    correct_answer = int(correct_answer_info.iloc[0])
    
    p_correct = row[f'선지정답률_{correct_answer}']
    p_correct = np.clip(p_correct, 0.01, 0.99)
    difficulty_b = -np.log(p_correct / (1 - p_correct))
    
    distractor_ps = [row[f'선지정답률_{i}'] for i in range(1, 6) if i != correct_answer]
    p_incorrect = 1 - p_correct
    
    if p_incorrect < 0.01 or not distractor_ps:
        discrimination_a = 0.5
    else:
        max_distractor_p = max(distractor_ps)
        attractiveness = max_distractor_p / p_incorrect if p_incorrect > 0 else 0
        discrimination_a = np.clip(0.5 + 2.5 * (attractiveness - 0.2), 0.3, 2.5)

    new_irt_values.append({
        '문제id': problem_id,
        'irt_difficulty_b': round(difficulty_b, 4),
        'irt_discrimination_a': round(discrimination_a, 4)
    })

df_new_irt = pd.DataFrame(new_irt_values)
print(f"{len(df_new_irt)}개 문항에 대한 새로운 IRT 값을 계산했습니다.")

# 5. 마스터 DB에 새로운 IRT 값 갱신
# '문제id'를 기준으로 새로운 값을 업데이트 (merge 사용)
master_df.set_index('문제id', inplace=True)
df_new_irt.set_index('문제id', inplace=True)
master_df.update(df_new_irt)
master_df.reset_index(inplace=True)

# 갱신된 마스터 DB 저장
master_df.to_excel(MASTER_DB_PATH, index=False)

print("\n--- IRT 값 갱신 완료 ---")
print(f"마스터 DB 파일 '{MASTER_DB_PATH}'이(가) 새로운 값으로 갱신되었습니다.")
print("갱신된 데이터 샘플:")
print(master_df[master_df['문제id'].isin(df_new_irt.index)][['문제id', 'irt_difficulty_b', 'irt_discrimination_a']].head())