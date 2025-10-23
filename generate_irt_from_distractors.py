import pandas as pd
import numpy as np
import os

# --- 설정 ---
# (필수) 선지별 정답률이 포함된 원본 Excel 파일 경로
DB_FILE_PATH = r"./data/db.xlsx"

# (필수) IRT 컬럼이 추가된 새 파일을 저장할 경로
OUTPUT_FILE_PATH = r"db_with_irt_from_distractors.xlsx"
# ------------

print(f"선지별 정답률 기반 IRT 모의 값 생성 시작...")
print(f"입력 파일: {DB_FILE_PATH}")

if not os.path.exists(DB_FILE_PATH):
    print(f"오류: DB 파일({DB_FILE_PATH})을 찾을 수 없습니다.")
    print("스크립트 상단의 DB_FILE_PATH 변수를 파일의 실제 위치로 수정하세요.")
else:
    try:
        df = pd.read_excel(DB_FILE_PATH)

        print(f"DB 읽기 완료: {len(df)}행")

        # 분석에 필요한 컬럼 확인
        required_cols = ['정답', '선지정답률_1', '선지정답률_2', '선지정답률_3', '선지정답률_4', '선지정답률_5']
        if not all(col in df.columns for col in required_cols):
            print(f"오류: {required_cols} 중 하나 이상의 컬럼이 파일에 없습니다.")
            print(f"현재 파일의 컬럼: {df.columns.tolist()}")
        else:
            irt_values = []
            
            for index, row in df.iterrows():
                try:
                    correct_answer = int(row['정답'])
                except (ValueError, TypeError):
                    # 정답 값이 숫자가 아니거나 비어있는 경우
                    irt_values.append({'irt_difficulty_b': np.nan, 'irt_discrimination_a': np.nan})
                    continue

                if not (1 <= correct_answer <= 5):
                    irt_values.append({'irt_difficulty_b': np.nan, 'irt_discrimination_a': np.nan})
                    continue

                # 정답률 (p_correct)
                p_correct = row[f'선지정답률_{correct_answer}']
                
                # p_correct가 0 또는 1이면 logit 변환 시 무한대가 되므로, 0.01 ~ 0.99 사이로 제한
                p_correct = np.clip(p_correct, 0.01, 0.99)

                # 1. 난이도(b) 계산 (로짓 변환)
                difficulty_b = -np.log(p_correct / (1 - p_correct))
                
                # 2. 변별도(a) 계산
                distractor_ps = [row[f'선지정답률_{i}'] for i in range(1, 6) if i != correct_answer]
                
                p_incorrect = 1 - p_correct
                
                if p_incorrect < 0.01 or not distractor_ps:
                    discrimination_a = 0.5 # 오답자가 거의 없으면 변별력 낮음
                else:
                    max_distractor_p = max(distractor_ps)
                    attractiveness = max_distractor_p / p_incorrect if p_incorrect > 0 else 0
                    discrimination_a = 0.5 + 2.5 * (attractiveness - 0.2)
                    discrimination_a = np.clip(discrimination_a, 0.3, 2.5) # 극단값 제한

                irt_values.append({
                    'irt_difficulty_b': round(difficulty_b, 4),
                    'irt_discrimination_a': round(discrimination_a, 4)
                })

            df_irt = pd.DataFrame(irt_values, index=df.index)
            
            df['irt_difficulty_b'] = df_irt['irt_difficulty_b']
            df['irt_discrimination_a'] = df_irt['irt_discrimination_a']
            
            df.to_excel(OUTPUT_FILE_PATH, index=False)
            
            print("\n--- 선지별 정답률 기반 IRT 모의 값 생성 완료 ---")
            print(f"생성된 파일: {OUTPUT_FILE_PATH}")
            print("\n생성된 데이터 샘플 (상위 5개):")
            print(df[['문제id', '정답', '선지정답률_1', 'irt_difficulty_b', 'irt_discrimination_a']].head())
            print(f"\n이제 {OUTPUT_FILE_PATH} 파일을 맞춤형 시험지 생성에 사용하세요.")

    except Exception as e:
        print(f"스크립트 실행 중 오류 발생: {e}")