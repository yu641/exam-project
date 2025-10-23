import os
import logging
import pandas as pd
import openpyxl
import re
import numpy as np
from datetime import datetime, timezone, timedelta
import traceback
import pythoncom

# --- 로깅 및 시간 설정 ---
log_gen = logging.getLogger("exam_generator")
log_grade = logging.getLogger("exam_grader")
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
KST = timezone(timedelta(hours=9), name="KST")

# --- KST 및 Exam ID 생성 유틸 ---
def make_exam_id(student_id: str, now: datetime, exam_count: int) -> str:
    """
    시험 ID를 학생 ID, 날짜, 시험 횟수(N차)로 구성합니다.
    (예: S001_20251023_1)
    """
    date_str = now.strftime("%Y%m%d")
    return f"{student_id}_{date_str}_{exam_count}"

# ---------- Word 유틸 ----------
def cm_to_pt(v: float) -> float:
    return (72.0 / 2.54) * float(v)

def start_word(visible=True):
    try:
        pythoncom.CoInitialize()
        import win32com.client as win32
        word = win32.gencache.EnsureDispatch("Word.Application")
    except Exception:
        log_gen.error("Word Application 로드 실패. win32com 캐시를 정리하거나 Office 설치 복구가 필요할 수 있습니다.")
        log_gen.error(traceback.format_exc())
        return None, None
    word.Visible = visible
    wd = win32.constants
    word.Options.PasteFormatBetweenDocuments = wd.wdKeepSourceFormatting
    word.Options.PasteFormatBetweenStyledDocuments = wd.wdKeepSourceFormatting
    word.Options.PasteAdjustTableFormatting = True
    return word, wd

def ensure_style(doc, name, wd, font="학교안심 바른바탕 R", size=11, bold=False, align=None, space_after=6):
    try:
        s = doc.Styles(name)
    except Exception:
        s = doc.Styles.Add(Name=name, Type=wd.wdStyleTypeParagraph)
    s.Font.Name = font; s.Font.Size = size; s.Font.Bold = bold
    if align is not None:
        s.ParagraphFormat.Alignment = align
    s.ParagraphFormat.SpaceAfter = space_after
    return s

def end_range(doc, wd):
    rng = doc.Range()
    rng.Collapse(wd.wdCollapseEnd)
    return rng

def insert_paragraph(doc, wd, text, style_name=None):
    rng = end_range(doc, wd)
    rng.InsertAfter(text)
    rng.InsertParagraphAfter()
    if style_name:
        doc.Paragraphs.Last.Range.Style = style_name

def insert_docx_with_source_format(doc, wd, path, word_app):
    if not os.path.exists(path):
        log_gen.warning(f"파일 없음: {path}")
        return False
    src = None
    try:
        src = word_app.Documents.Open(
            FileName=os.path.abspath(path), ReadOnly=True,
            AddToRecentFiles=False, Visible=False
        )
        src.Content.Copy()
        dest = end_range(doc, wd)
        dest.PasteAndFormat(wd.wdFormatOriginalFormatting)
        return True
    except Exception:
        log_gen.error(f"삽입 실패({path}):")
        log_gen.error(traceback.format_exc())
        return False
    finally:
        if src is not None:
            try: src.Close(False)
            except Exception: pass

def set_two_columns_current_section(doc, wd, line_between=False, count=2):
    try:
        last_para = doc.Paragraphs.Last
        if last_para.Range.Text.strip() != "":
            last_para.Range.InsertParagraphAfter()
        rng_for_break = end_range(doc, wd)
        rng_for_break.InsertBreak(wd.wdSectionBreakContinuous)
        sec = doc.Sections.Last
        sec.PageSetup.TextColumns.SetCount(count)
        sec.PageSetup.TextColumns.LineBetween = bool(line_between)
    except Exception:
        log_gen.error("2단 나누기 설정 중 오류 발생:")
        log_gen.error(traceback.format_exc())

# ---------- Word 문서 생성 내부 함수 (안정화) ----------
def _create_word_document(tasks, title, subtitle, student_name, two_columns, output_dir, student_id):
    word, wd = start_word(visible=True)
    if not word: return None
    
    doc = None
    out_path = None
    try:
        doc = word.Documents.Add()
        ps = doc.Sections(1).PageSetup
        ps.TopMargin = cm_to_pt(1.5); ps.BottomMargin = cm_to_pt(1.5)
        ps.LeftMargin = cm_to_pt(1.5); ps.RightMargin = cm_to_pt(1.5)

        ensure_style(doc, "제목", wd, font="학교안심 바른바탕 B", size=24, bold=True, align=wd.wdAlignParagraphCenter, space_after=12)
        ensure_style(doc, "소제목", wd, font="학교안심 바른바탕 R", size=16, align=wd.wdAlignParagraphLeft, space_after=6)
        ensure_style(doc, "문항", wd, font="학교안심 바른바탕 R", size=11, align=wd.wdAlignParagraphLeft, space_after=6)

        insert_paragraph(doc, wd, title, style_name="제목")
        if subtitle: insert_paragraph(doc, wd, subtitle, style_name="소제목")
        insert_paragraph(doc, wd, f"학년: ____  반: ____  번호: ____  이름: {student_name}")
        end_range(doc, wd).InsertParagraphAfter()

        if two_columns:
            set_two_columns_current_section(doc, wd, line_between=False, count=2)

        passage_number, problem_number_in_passage = 1, 1
        for tag, path in tasks:
            if tag.startswith("지문 "):
                if problem_number_in_passage > 1: passage_number += 1
                insert_paragraph(doc, wd, f"[{passage_number}]", style_name="문항")
                insert_docx_with_source_format(doc, wd, path, word_app=word)
                problem_number_in_passage = 1
            else:
                insert_paragraph(doc, wd, f"{passage_number}-{problem_number_in_passage})", style_name="문항")
                insert_docx_with_source_format(doc, wd, path, word_app=word)
                problem_number_in_passage += 1

        out_name = f"{title.replace(' ', '_')}_{student_id}.docx"
        out_path = os.path.abspath(os.path.join(output_dir, out_name))
        doc.SaveAs(out_path)
        log_gen.info(f"시험지 생성 완료: {out_path}")
    except Exception:
        log_gen.error("Word 문서 생성 프로세스에서 예외 발생:")
        log_gen.error(traceback.format_exc())
        out_path = None
    finally:
        if doc:
            try: doc.Close(False)
            except Exception: pass
        if word:
            try: word.Quit()
            except Exception: pass
    return out_path

# ---------- 1. 첫 사용자용 시험지 생성 (랜덤 7지문) ----------
def generate_exam_7_passages_from_db(
    db_path: str, base_dir: str, title: str, subtitle: str | None = None,
    two_columns: bool = True, output_dir: str = "./output",
    student_id: str = "S000", student_name: str = "학생"
):
    log_gen.info(f"랜덤 시험지 생성 시작 (학생: {student_name}, ID: {student_id})")
    os.makedirs(output_dir, exist_ok=True)

    try:
        df_db = pd.read_excel(db_path, sheet_name=0)
    except Exception as e:
        log_gen.error(f"DB 읽기 실패: {e}")
        return None, None, None
    
    need = {"지문id","문제id","지문유형","문제유형","정답","과목"}
    if not all(c in df_db.columns for c in need):
        log_gen.error(f"DB에 필요한 컬럼 부족: {set(need) - set(df_db.columns)}")
        return None, None, None

    df_passages = df_db.loc[:, ["지문id","지문유형"]].dropna().drop_duplicates(subset=["지문id"])
    if df_passages.empty:
        log_gen.error("DB에 지문 후보가 없습니다.")
        return None, None, None

    categories = {
        "독서": [["인문","주제통합","예술"], ["과학","기술","과학기술","과학·기술"], ["사회"]],
        "문학": [["현대시","현대시*"], ["현대소설"], ["고전시가","고전시가*"], ["고전소설"]],
    }
    rng = np.random.default_rng()
    selected_passage_rows = []
    for group_list in categories.values():
        for group in group_list:
            hits = df_passages[df_passages["지문유형"].astype(str).str.contains('|'.join(group), na=False)]
            if not hits.empty:
                selected_passage_rows.append(hits.sample(1, random_state=int(rng.integers(0, 1_000_000))).iloc[0])

    if not selected_passage_rows:
        log_gen.error("카테고리에 맞는 지문을 찾을 수 없습니다.")
        return None, None, None
    selected_passages = pd.DataFrame(selected_passage_rows).drop_duplicates(subset=["지문id"]).reset_index(drop=True)

    tasks, selected_records = [], []
    for _, row in selected_passages.iterrows():
        pid = row["지문id"]
        p_path = os.path.join(base_dir, "지문", f"{pid}.docx")
        if os.path.exists(p_path):
            tasks.append((f"지문 {pid}", p_path))
        else:
            log_gen.warning(f"지문 파일 없음: {p_path}")

        rel = df_db[df_db["지문id"] == pid]
        for _, q in rel.iterrows():
            qid = str(q["문제id"])
            q_path = os.path.join(base_dir, "문제", f"{qid}.docx")
            if os.path.exists(q_path):
                tasks.append((f"문제 {qid}", q_path))
                selected_records.append({"problem_id": qid, "answer": q.get("정답", ""),
                                         "subject": q.get("과목", ""), "problem_type": q.get("문제유형", "")})
            else:
                log_gen.warning(f"문제 파일 없음: {q_path}")

    if not tasks:
        log_gen.error("삽입할 유효한 파일(지문/문제)이 없습니다.")
        return None, None, None

    out_path = _create_word_document(tasks, title, subtitle, student_name, two_columns, output_dir, student_id)
    
    now = datetime.now(KST)
    
    count = 0
    try:
        for fname in os.listdir(output_dir):
            if (fname.startswith(f"{student_id}_") and 
                fname.endswith(".xlsx") and
                "_graded.xlsx" not in fname and
                "_result.xlsx" not in fname and
                not fname.startswith("user_weakness_")):
                count += 1
    except FileNotFoundError:
        pass
    exam_count = count + 1
    
    exam_id = make_exam_id(student_id=student_id, now=now, exam_count=exam_count)
    
    meta_xlsx_path = os.path.abspath(os.path.join(output_dir, f"{exam_id}.xlsx"))
    meta_df = pd.DataFrame([{"exam_id": exam_id, "student_id": student_id,
                             "student_name": student_name, "exam_name": title,
                             "timestamp": now.isoformat(timespec="seconds")}])

    with pd.ExcelWriter(meta_xlsx_path, engine="openpyxl") as writer:
        pd.DataFrame(selected_records).to_excel(writer, index=False, sheet_name="selected_problems")
        meta_df.to_excel(writer, index=False, sheet_name="meta")
    log_gen.info(f"메타 파일 저장 완료: {meta_xlsx_path}")

    return out_path, meta_xlsx_path, exam_id

# ---------- 2. 맞춤형 시험지 생성 (IRT + 취약점) ----------
def generate_exam_irt_weakness(
    db_path: str, base_dir: str, user_weakness_path: str, user_theta: float,
    title: str, subtitle: str | None = None, num_passages: int = 7,
    num_problems_per_passage: int = 4, weak_passage_target_prop: float = 0.6,
    weak_problem_boost: float = 1.5, two_columns: bool = True,
    output_dir: str = "./output", student_id: str = "S000", student_name: str = "학생"
):
    log_gen.info(f"맞춤형(IRT) 시험지 생성 시작 (학생: {student_name}, ID: {student_id}, Theta: {user_theta})")
    os.makedirs(output_dir, exist_ok=True)
    rng = np.random.default_rng()

    try:
        df_db = pd.read_excel(db_path, sheet_name=0)
    except Exception as e:
        log_gen.error(f"DB 읽기 실패: {e}")
        return None, None, None
    
    need = {"지문id","문제id","지문유형","문제유형","정답","과목", "irt_difficulty_b", "irt_discrimination_a"}
    if not all(c in df_db.columns for c in need):
        log_gen.error(f"DB에 필요한 IRT 컬럼 부족: {set(need) - set(df_db.columns)}")
        return None, None, None

    weak_passage_set, weak_problem_set = set(), set()
    if not os.path.exists(user_weakness_path):
        log_gen.warning(f"취약점 파일({user_weakness_path})을 찾을 수 없음. 랜덤 선택으로 진행.")
        weak_passage_target_prop = 0.0
    else:
        try:
            weak_passage_set = set(pd.read_excel(user_weakness_path, sheet_name="weak_passages")['지문유형코드'])
            weak_problem_set = set(pd.read_excel(user_weakness_path, sheet_name="weak_problems")['문제유형코드'])
            log_gen.info(f"취약 지문({len(weak_passage_set)}개), 취약 문제({len(weak_problem_set)}개) 로드")
        except Exception as e:
            log_gen.error(f"사용자 취약점 파일 읽기 실패: {e}")
            return None, None, None

    df_passages = df_db.loc[:, ["지문id","지문유형"]].dropna().drop_duplicates(subset=["지문id"])
    if df_passages.empty:
        log_gen.error("DB에 지문 후보가 없습니다.")
        return None, None, None
    
    df_passages['is_weak'] = df_passages['지문유형'].isin(weak_passage_set)
    n_weak_target = int(num_passages * weak_passage_target_prop)
    df_weak_available = df_passages[df_passages['is_weak']]
    df_other_available = df_passages[~df_passages['is_weak']]
    n_weak = min(n_weak_target, len(df_weak_available))
    n_other = num_passages - n_weak
    if n_other > len(df_other_available):
        log_gen.warning(f"비취약 지문 수 부족 (필요: {n_other}, 가능: {len(df_other_available)})")
        n_other = len(df_other_available)
        n_weak = min(num_passages - n_other, len(df_weak_available))
    
    df_weak_selected = df_weak_available.sample(n_weak, random_state=rng) if n_weak > 0 else pd.DataFrame()
    df_other_selected = df_other_available.sample(n_other, random_state=rng) if n_other > 0 else pd.DataFrame()
    selected_passages = pd.concat([df_weak_selected, df_other_selected]).sample(frac=1, random_state=rng)
    log_gen.info(f"지문 선택: 총 {len(selected_passages)}개 (취약 {n_weak}개, 일반 {n_other}개)")

    tasks, selected_records = [], []
    for _, p_row in selected_passages.iterrows():
        pid = p_row["지문id"]
        p_path = os.path.join(base_dir, "지문", f"{pid}.docx")
        if os.path.exists(p_path):
            tasks.append((f"지문 {pid}", p_path))
        else:
            log_gen.warning(f"지문 파일 없음: {p_path} (스킵)")
            continue

        related_problems = df_db[df_db["지문id"] == pid].copy()
        if related_problems.empty: continue
            
        related_problems['is_weak'] = related_problems['문제유형'].isin(weak_problem_set)
        info_score = related_problems['irt_discrimination_a'] / (1.0 + np.abs(related_problems['irt_difficulty_b'] - user_theta))
        weak_bonus = np.where(related_problems['is_weak'], weak_problem_boost, 1.0)
        related_problems['final_score'] = info_score * weak_bonus
        selected_problems = related_problems.nlargest(num_problems_per_passage, 'final_score')

        for _, q in selected_problems.iterrows():
            qid = str(q["문제id"])
            q_path = os.path.join(base_dir, "문제", f"{qid}.docx")
            if os.path.exists(q_path):
                tasks.append((f"문제 {qid}", q_path))
                selected_records.append({"problem_id": qid, "answer": q.get("정답", ""),
                                         "subject": q.get("과목", ""), "problem_type": q.get("문제유형", "")})
            else:
                log_gen.warning(f"문제 파일 없음: {q_path}")

    if not tasks:
        log_gen.error("삽입할 유효한 파일(지문/문제)이 없습니다.")
        return None, None, None

    out_path = _create_word_document(tasks, title, subtitle, student_name, two_columns, output_dir, student_id)
    
    now = datetime.now(KST)
    
    count = 0
    try:
        for fname in os.listdir(output_dir):
            if (fname.startswith(f"{student_id}_") and 
                fname.endswith(".xlsx") and
                "_graded.xlsx" not in fname and
                "_result.xlsx" not in fname and
                not fname.startswith("user_weakness_")):
                count += 1
    except FileNotFoundError:
        pass
    exam_count = count + 1
    
    exam_id = make_exam_id(student_id=student_id, now=now, exam_count=exam_count)
    
    meta_xlsx_path = os.path.abspath(os.path.join(output_dir, f"{exam_id}.xlsx"))
    meta_df = pd.DataFrame([{"exam_id": exam_id, "student_id": student_id,
                             "student_name": student_name, "exam_name": title,
                             "timestamp": now.isoformat(timespec="seconds"),
                             "user_theta": user_theta}])

    with pd.ExcelWriter(meta_xlsx_path, engine="openpyxl") as writer:
        pd.DataFrame(selected_records).to_excel(writer, index=False, sheet_name="selected_problems")
        meta_df.to_excel(writer, index=False, sheet_name="meta")
    log_gen.info(f"메타 파일 저장 완료: {meta_xlsx_path}")

    return out_path, meta_xlsx_path, exam_id

# ---------- 3. 채점 및 분석 함수 ----------
def normalize_answer_num(x) -> str:
    if x is None: return ""
    s = str(x).strip().upper()
    s = re.sub(r'[\s\(\)]', '', s)
    alpha_to_num = {"A":"1", "B":"2", "C":"3", "D":"4", "E":"5"}
    return alpha_to_num.get(s, s if s in {"1","2","3","4","5"} else "")

def _read_answers_excel_first_col(path: str, expected_len: int) -> list[str]:
    try:
        df = pd.read_excel(path, sheet_name=0, header=None)
    except Exception as e:
        log_grade.warning(f"답안 엑셀 읽기 실패 ({e}). 빈 답안으로 처리.")
        return [""] * expected_len
    if df.empty: return [""] * expected_len
    col0 = df.iloc[:, 0].tolist()
    norm = [normalize_answer_num(v) for v in col0]
    if len(norm) < expected_len: norm += [""] * (expected_len - len(norm))
    else: norm = norm[:expected_len]
    return norm

def grade_exam(
    exam_xlsx_path: str,
    answers: dict[str, str] | None = None,
    interactive: bool = True,
    output_dir: str | None = None,
    answers_xlsx_path: str | None = None,
) -> dict:
    if not os.path.exists(exam_xlsx_path):
        log_grade.error(f"파일을 찾을 수 없습니다: {exam_xlsx_path}")
        return {}

    try:
        sel  = pd.read_excel(exam_xlsx_path, sheet_name="selected_problems")
        meta = pd.read_excel(exam_xlsx_path, sheet_name="meta")
    except Exception as e:
         log_grade.error(f"파일 시트 읽기 실패. 'selected_problems'/'meta' 시트 필요. ({e})")
         return {}

    if meta.empty:
        log_grade.error("meta 시트가 비어있습니다.")
        return {}
    
    meta_row = meta.iloc[0]
    student_id   = str(meta_row.get("student_id", "UNKNOWN_ID"))
    student_name = str(meta_row.get("student_name", "UNKNOWN_NAME"))
    exam_id      = str(meta_row.get("exam_id", ""))
    exam_name    = str(meta_row.get("exam_name", ""))
    created_ts   = str(meta_row.get("timestamp", ""))
    
    log_grade.info(f"채점 시작 - 시험 ID: {exam_id}, 학생: {student_name} ({student_id})")

    need_cols = {"problem_id", "answer", "subject", "problem_type"}
    if not all(c in sel.columns for c in need_cols):
        log_grade.error(f"selected_problems 시트에 분석용 필수 컬럼이 없습니다: {need_cols - set(sel.columns)}")
        return {}

    sel["problem_id"] = sel["problem_id"].astype(str)
    sel["answer_num"] = sel["answer"].apply(normalize_answer_num)

    submitted = []
    if answers_xlsx_path and os.path.exists(answers_xlsx_path):
        seq = _read_answers_excel_first_col(answers_xlsx_path, expected_len=len(sel))
        submitted = [{"problem_id": pid, "student_answer_num": a} for pid, a in zip(sel["problem_id"].tolist(), seq)]
    elif interactive:
        print("\n정답을 입력하세요. (1~5만 허용)")
        for _, r in sel.iterrows():
            raw = input(f"문항 ID {r['problem_id']} [{r['subject']}/{r['problem_type']}] : ").strip()
            submitted.append({"problem_id": r['problem_id'], "student_answer_num": normalize_answer_num(raw)})
    else:
        submitted = [{"problem_id": pid, "student_answer_num": ""} for pid in sel["problem_id"]]

    ans_df = pd.DataFrame(submitted)
    g = sel.merge(ans_df, on="problem_id", how="left")
    g["student_answer_num"] = g["student_answer_num"].fillna("")
    g["is_correct"] = np.where(g["answer_num"] == g["student_answer_num"], 1, 0)
    g.loc[g["answer_num"] == "", "is_correct"] = 0

    total, correct = len(g), int(g["is_correct"].sum())
    score = round(correct / total * 100, 2) if total else 0.0

    if output_dir is None: output_dir = os.path.dirname(exam_xlsx_path)
    os.makedirs(output_dir, exist_ok=True)
    base = os.path.basename(exam_xlsx_path).replace(".xlsx", "")
    graded_path = os.path.join(output_dir, f"{base}_graded.xlsx")
    result_path = os.path.join(output_dir, f"{base}_result.xlsx")
    submitted_at = datetime.now(KST).isoformat(timespec="seconds")

    summary_sheet = pd.DataFrame([{"exam_id": exam_id, "exam_name": exam_name, "created_at": created_ts,
        "submitted_at": submitted_at, "student_id": student_id, "student_name": student_name,
        "total": total, "correct": correct, "score": score}])
    
    by_subject = (g.groupby("subject")["is_correct"].agg(total="size", correct="sum").reset_index())
    by_subject["accuracy(%)"] = (by_subject["correct"]/by_subject["total"]*100).round(2)

    with pd.ExcelWriter(graded_path, engine="openpyxl") as writer:
        meta.to_excel(writer, index=False, sheet_name="meta")
        sel.to_excel(writer, index=False, sheet_name="selected_problems")
        g[['problem_id', 'student_answer_num']].assign(submitted_at=submitted_at, student_id=student_id, student_name=student_name).to_excel(writer, index=False, sheet_name="answers")
        
        grading_cols = list(need_cols | {"answer_num", "student_answer_num", "is_correct"})
        g[grading_cols].to_excel(writer, index=False, sheet_name="grading")

        summary_sheet.to_excel(writer, index=False, sheet_name="summary")
        if not by_subject.empty:
            by_subject.to_excel(writer, index=False, sheet_name="summary_by_subject")

    with pd.ExcelWriter(result_path, engine="openpyxl") as writer:
        summary_sheet.to_excel(writer, index=False, sheet_name="result")
        g.rename(columns={"answer_num": "answer", "student_answer_num": "my_answers"})[['problem_id', 'answer', 'my_answers', 'subject', 'problem_type']].to_excel(writer, index=False, sheet_name="details")
        meta.to_excel(writer, index=False, sheet_name="meta")

    log_grade.info(f"채점 완료 (상세: {graded_path}, 간단: {result_path})")
    log_grade.info(f"총 {total}문항, 정답 {correct}개, 점수 {score}점")

    return {"graded_path": graded_path, "result_path": result_path, "exam_id": exam_id, 
            "student_id": student_id, "student_name": student_name, "total": total, 
            "correct": correct, "score": score}

# 취약점 갱신 함수
def analyze_weakness_from_graded_file(
    graded_xlsx_path: str, output_dir: str,
    passage_threshold: float = 70.0, problem_threshold: float = 60.0
) -> str | None:
    """
    채점 완료된 ..._graded.xlsx 파일을 분석하여 취약점 파일(user_weakness_...xlsx)을
    누적 갱신합니다.
    """
    if not os.path.exists(graded_xlsx_path):
        log_grade.error(f"취약점 분석 실패: 채점 파일을 찾을 수 없음 ({graded_xlsx_path})")
        return None
    
    try:
        # --- 1. Get Student ID and define output path ---
        student_id = str(pd.read_excel(graded_xlsx_path, sheet_name="meta").iloc[0]["student_id"])
        log_grade.info(f"{student_id} 학생 취약점 분석 시작...")
        output_path = os.path.join(output_dir, f"user_weakness_{student_id}.xlsx")

        # --- 2. Get new analysis data from the graded file ---
        
        # 2a. Get exam metadata (ID, time) from summary sheet
        df_summary = pd.read_excel(graded_xlsx_path, sheet_name="summary")
        summary_row = df_summary.iloc[0]
        exam_id = summary_row.get("exam_id", "UNKNOWN_EXAM")
        
        # submitted_at에서 날짜(YYYY-MM-DD)만 추출
        submitted_at_full = summary_row.get("submitted_at", datetime.now(KST).isoformat(timespec="seconds"))
        submitted_at_date_only = submitted_at_full.split('T')[0]
        
        try:
            # exam_id (예: S001_20251023_1)에서 횟수(1)를 추출
            exam_count = int(exam_id.split('_')[-1])
        except (IndexError, ValueError):
            exam_count = 0 # Fallback
        
        # 로그 메시지에도 날짜 포함
        log_grade.info(f"분석 대상 시험: {exam_id} (횟수: {exam_count}, 제출: {submitted_at_date_only})")

        # 2b. Find weak passages
        df_subject = pd.read_excel(graded_xlsx_path, sheet_name="summary_by_subject")
        df_subject['accuracy'] = (df_subject['correct'] / df_subject['total']) * 100
        weak_passages_found = df_subject[df_subject['accuracy'] < passage_threshold]
        
        # 2c. Find weak problems
        df_grading = pd.read_excel(graded_xlsx_path, sheet_name="grading")
        df_problem_analysis = df_grading.groupby("problem_type")['is_correct'].agg(
            total='count', correct='sum').reset_index()
        df_problem_analysis['accuracy'] = (df_problem_analysis['correct'] / df_problem_analysis['total']) * 100
        weak_problems_found = df_problem_analysis[df_problem_analysis['accuracy'] < problem_threshold]
        
        # 2d. Format new weakness dataframes
        df_new_passages = pd.DataFrame({"지문유형코드": weak_passages_found['subject']})
        df_new_passages['exam_id'] = exam_id
        df_new_passages['submitted_at'] = submitted_at_date_only
        df_new_passages['exam_count'] = exam_count
        
        df_new_problems = pd.DataFrame({"문제유형코드": weak_problems_found['problem_type']})
        df_new_problems['exam_id'] = exam_id
        df_new_problems['submitted_at'] = submitted_at_date_only 

        # --- 3. Load old weakness data (if exists) ---
        df_old_passages = pd.DataFrame()
        df_old_problems = pd.DataFrame()
        if os.path.exists(output_path):
            try:
                df_old_passages = pd.read_excel(output_path, sheet_name="weak_passages")
                df_old_problems = pd.read_excel(output_path, sheet_name="weak_problems")
                log_grade.info(f"기존 취약점 파일 로드: {output_path}")
            except Exception as e:
                log_grade.warning(f"기존 취약점 파일({output_path}) 읽기 실패. 새 파일로 덮어씁니다. (오류: {e})")

        # --- 4. Combine old and new data ---
        final_passages_df = pd.concat([df_old_passages, df_new_passages]).drop_duplicates()
        final_problems_df = pd.concat([df_old_problems, df_new_problems]).drop_duplicates()

        # --- 5. Save combined data ---
        os.makedirs(output_dir, exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            final_passages_df.to_excel(writer, index=False, sheet_name="weak_passages")
            final_problems_df.to_excel(writer, index=False, sheet_name="weak_problems")

        log_grade.info(f"취약점 분석 완료 (누적 갱신): {output_path}")
        return output_path
    
    except Exception as e:
        log_grade.error(f"취약점 분석 중 오류 발생: {e}")
        log_grade.error(traceback.format_exc())
        return None