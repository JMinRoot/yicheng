# -*- coding: utf-8 -*-
"""
考试酷(examcoo.com) 成绩批量下载脚本
- 自动获取pid、班级名称、考试名称
- 分页获取全部成绩数据
- 保留每题得分
- 多答卷合并逻辑：主观题(满分100)和其他题型(满分100)分别取最高分后合并
- 不同班级的成绩保存到不同的Excel文件，每场考试一个Sheet
"""

import requests
import re
import json
import os
import time
import urllib3
import pandas as pd

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 禁用SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ===================== 文件锁处理 =====================
import shutil
MAX_FILE_RETRY = 10       # 最大重试次数
FILE_RETRY_INTERVAL = 3   # 每次重试间隔（秒）


def is_file_locked(filepath):
    """检测文件是否被其他程序锁定（如Excel正在打开）"""
    if not os.path.exists(filepath):
        return False
    try:
        # Windows下尝试以独占模式打开，失败说明被锁定
        fd = os.open(filepath, os.O_RDWR | os.O_EXCL)
        os.close(fd)
        return False
    except OSError:
        return True


def safe_read_excel(filepath, sheet_name=0, **kwargs):
    """
    安全读取Excel，遇到文件锁定自动重试。
    返回 DataFrame（sheet_name=None时返回dict）或 None（彻底失败）
    """
    filename = os.path.basename(filepath)
    for attempt in range(1, MAX_FILE_RETRY + 1):
        try:
            return pd.read_excel(filepath, sheet_name=sheet_name, **kwargs)
        except PermissionError as e:
            if attempt < MAX_FILE_RETRY:
                print(f"  [警告] 文件 '{filename}' 被占用，{FILE_RETRY_INTERVAL}秒后重试... ({attempt}/{MAX_FILE_RETRY})")
                time.sleep(FILE_RETRY_INTERVAL)
            else:
                print(f"  [错误] 文件 '{filename}' 持续被占用，请关闭该文件后重试！")
                print(f"         路径: {filepath}")
                return None
        except Exception as e:
            print(f"  [错误] 读取文件 '{filename}' 失败: {e}")
            return None


def safe_save_workbook(wb, filepath):
    """
    安全保存Workbook，遇到文件锁定自动重试（先保存到临时文件再替换）。
    成功返回True，失败返回False
    """
    filename = os.path.basename(filepath)
    tmp_path = filepath + ".tmp"
    for attempt in range(1, MAX_FILE_RETRY + 1):
        try:
            wb.save(tmp_path)
            # 替换原文件（如果原文件被锁定这里也会报错）
            try:
                os.replace(tmp_path, filepath)
            except PermissionError:
                # Windows上os.replace对已打开的文件也可能失败，尝试shutil
                shutil.move(tmp_path, filepath)
            return True
        except (PermissionError, OSError) as e:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
            if attempt < MAX_FILE_RETRY:
                print(f"  [警告] 无法写入 '{filename}'（文件可能已被打开），{FILE_RETRY_INTERVAL}秒后重试... ({attempt}/{MAX_FILE_RETRY})")
                time.sleep(FILE_RETRY_INTERVAL)
            else:
                print(f"  [错误] 无法写入 '{filename}'，请关闭该文件后重试！")
                print(f"         路径: {filepath}")
                return False
        except Exception as e:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
            print(f"  [错误] 保存文件 '{filename}' 失败: {e}")
            return False


def safe_excel_write(df_dict_or_list, filepath, engine="openpyxl"):
    """
    安全通过pd.ExcelWriter写入Excel，遇到文件锁定自动重试。
    df_dict_or_list: dict of {sheet_name: DataFrame} 或 list of (sheet_name, DataFrame)
    成功返回True，失败返回False
    """
    filename = os.path.basename(filepath)
    tmp_path = filepath + ".tmp"

    # 统一为 list 格式
    if isinstance(df_dict_or_list, dict):
        items = list(df_dict_or_list.items())
    else:
        items = df_dict_or_list

    for attempt in range(1, MAX_FILE_RETRY + 1):
        try:
            with pd.ExcelWriter(tmp_path, engine=engine) as writer:
                for sheet_name, df in items:
                    safe_sheet = safe_sheet_name(sheet_name) if isinstance(sheet_name, str) else str(sheet_name)
                    df.to_excel(writer, sheet_name=safe_sheet, index=False)

            # 替换原文件
            try:
                os.replace(tmp_path, filepath)
            except PermissionError:
                shutil.move(tmp_path, filepath)
            return True
        except (PermissionError, OSError) as e:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
            if attempt < MAX_FILE_RETRY:
                print(f"  [警告] 无法写入 '{filename}'（文件可能已被打开），{FILE_RETRY_INTERVAL}秒后重试... ({attempt}/{MAX_FILE_RETRY})")
                time.sleep(FILE_RETRY_INTERVAL)
            else:
                print(f"  [错误] 无法写入 '{filename}'，请关闭该文件后重试！")
                print(f"         路径: {filepath}")
                return False
        except Exception as e:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
            print(f"  [错误] 写入文件 '{filename}' 失败: {e}")
            return False

# ===================== 配置区域 =====================

# Cookie（从浏览器开发者工具复制，完整Cookie字符串）
COOKIE = "lastspl=dae698870e6af9d5deb518acf97bf32a; lastuid=15298143; autoUid=15298143; autoPassword=39b5eef9ec33a632bdce101960533acf; autoTimezone=28800; PHPSESSID=07ivfei3lr36tsjbl0skt044f4"

# 考试列表：每项填 cid(班级ID) 和 tid(考试ID)
# 从URL中提取：/class/score/paperanalysis/cid/{cid}/tid/{tid}/idpaging/0
EXAMS = [
    #头狼考试

    {"cid": "1100056", "tid": "5628408"},#第1节
    {"cid": "1100056", "tid": "5639097"},#第2节
    {"cid": "1100056", "tid": "5654168"},#第3节
    {"cid": "1100056", "tid": "5654167"},#第3-1节
    {"cid": "1100056", "tid": "5666173"},#第4节
    {"cid": "1100056", "tid": "5676874"},#第5节
    {"cid": "1100056", "tid": "5676875"},#第6节
    {"cid": "1100056", "tid": "5685392"},#第7节
    {"cid": "1100056", "tid": "5697851"},#第8节
    {"cid": "1100056", "tid": "5700569"},#第9节 
    {"cid": "1100056", "tid": "5700573"},#第10节


    #树人考试
    {"cid": "1099008", "tid": "5628406"},#第1节
    {"cid": "1099008", "tid": "5639096"},#第2节
    {"cid": "1099008", "tid": "5654169"},#第3节
    {"cid": "1099008", "tid": "5666172"},#第4节
    {"cid": "1099008", "tid": "5676872"},#第5节
    {"cid": "1099008", "tid": "5676873"},#第6节
    {"cid": "1099008", "tid": "5685391"},#第7节
    {"cid": "1099008", "tid": "5697850"},#第8节
    {"cid": "1099008", "tid": "5700559"},#第9节
    {"cid": "1099008", "tid": "5700566"},#第10节
]

# 输出目录（默认桌面）
# OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop")

#输出到指定的文件夹
OUTPUT_DIR = os.path.join("C:/Users/Administrator/Desktop/树人计划第六期", "成绩处理")

# 请求间隔（秒），避免被限流
REQUEST_DELAY = 0.5

# 最大分页步数（安全限制）
MAX_STEPS = 30

# ===================== 核心逻辑 =====================

BASE_URL = "https://www.examcoo.com"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0",
    "Cookie": COOKIE,
}

# 题型映射
TYPE_MAP = {
    "1": "单选",
    "2": "多选",
    "3": "判断",
    "4": "填空",
    "5": "主观",
}


def get_page_info(cid, tid):
    """
    从试卷分析HTML页面提取 pid、班级名称、考试名称
    """
    url = f"{BASE_URL}/class/score/paperanalysis/cid/{cid}/tid/{tid}/idpaging/0"
    resp = requests.get(url, headers=HEADERS, verify=False, timeout=30)
    html = resp.text

    # 提取 gPid
    pid = None
    m = re.search(r'var\s+gPid\s*=\s*"(\d+)"', html)
    if m:
        pid = m.group(1)

    # 提取班级名称（格式：班级名称／cid）
    # HTML中用 &#12288; 表示全角空格，用更灵活的匹配
    class_name = ""
    m = re.search(r'班.*?级[：:].*?<td>(.*?)(?:／|<)', html, re.DOTALL)
    if m:
        class_name = m.group(1).strip()

    # 提取考试名称（格式：考试名称／pid）
    exam_name = ""
    m = re.search(r'考试试卷[：:].*?<td>(.*?)(?:／|<)', html, re.DOTALL)
    if m:
        exam_name = m.group(1).strip()

    return pid, class_name, exam_name


def fetch_all_exam_data(cid, tid, pid):
    """
    分页获取全部考试数据
    - step=0 返回 paperData（题目结构）+ 首批 examerData
    - hasData=True 时继续获取下一页
    """
    all_paper_data = None
    all_examer_data = []
    step = 0

    while step < MAX_STEPS:
        post_data = {
            "pid": pid,
            "tid": tid,
            "cid": cid,
            "idpaging": "0",
            "step": str(step),
        }
        referer = f"{BASE_URL}/class/score/paperanalysis/cid/{cid}/tid/{tid}/idpaging/0"
        req_headers = {**HEADERS, "Referer": referer}

        try:
            resp = requests.post(
                f"{BASE_URL}/class/score/examcontent",
                headers=req_headers,
                data=post_data,
                verify=False,
                timeout=60,
            )
            result = resp.json()
        except Exception as e:
            print(f"  [错误] 第{step}页请求失败: {e}")
            break

        # 检查返回值是否为对象
        if not isinstance(result, dict):
            print(f"  [错误] 第{step}页返回异常: {result}")
            break

        # step=0 时保存 paperData
        if step == 0:
            all_paper_data = result.get("paperData", [])

        examer_data = result.get("examerData", [])
        all_examer_data.extend(examer_data)
        print(f"  第{step}页: 获取 {len(examer_data)} 条记录")

        # 检查是否还有更多数据
        has_data = result.get("hasData")
        if not has_data or has_data == 0 or has_data == "0" or has_data is False:
            break

        step += 1
        time.sleep(REQUEST_DELAY)

    return all_paper_data, all_examer_data


def parse_questions(paper_data):
    """
    解析试卷结构，返回题目列表
    - id 以 's' 开头的是具体题目
    - id 以 'b' 开头的是大题分组（跳过）
    """
    questions = []
    q_num = 0
    for item in paper_data:
        qid = item.get("id", "")
        if not qid.startswith("s"):
            continue
        q_num += 1
        type_code = qid.split("_")[0][1:]  # 提取 s 后面的数字
        q_type = TYPE_MAP.get(type_code, f"题型{type_code}")
        questions.append({
            "id": qid,
            "num": q_num,
            "type": q_type,
            "points": float(item.get("p", 0)),
        })
    return questions


def parse_student_scores(examer_data, questions):
    """
    解析学员成绩数据
    - 提取每题得分（从 mark_content 中获取 s 字段）
    - 将题目分为主观题组(满分约100) 和其他题型组(满分约100)
    - 同一学员多次考试时：
      * 主观题组取主观题总分最高的那次答卷
      * 其他题型组取其他题型总分最高的那次答卷
      * 两组最佳答卷合并，确保主观题和其他题都能拿到各自最高分
    """
    # 题目 id -> 索引映射
    q_id_set = {q["id"] for q in questions}

    # 划分主观题 vs 其他题型
    subjective_qids = {q["id"] for q in questions if q["type"] == "主观"}
    other_qids = q_id_set - subjective_qids

    print(f"  题型划分: 主观题 {len(subjective_qids)} 道, 其他题型 {len(other_qids)} 道")

    # 原始记录列表（每个学员可能有多条）
    raw_records = []  # list of {uid, student_id, name, totalscore, q_scores}

    for item in examer_data:
        uid = str(item.get("uid", ""))
        if not uid:
            continue

        # 确定学员姓名
        name = item.get("userremark", "")
        if not name or name == "None":
            name = item.get("rsstudentname", "")
        if not name or name == "None":
            name = item.get("displayname", "")
        if not name or name == "None":

            name = f"用户{uid}"

        # 学号（如果有的话）
        student_id = item.get("rsstudentid", "")
        if not student_id or student_id == "None":
            student_id = uid

        total = float(item.get("totalscore", 0))

        # 解析每题得分
        mark_content_str = item.get("mark_content", "[]")
        try:
            mark_content = json.loads(mark_content_str)
        except (json.JSONDecodeError, TypeError):
            mark_content = []

        q_scores = {}
        for mq in mark_content:
            qid = mq.get("id", "")
            if qid not in q_id_set:
                continue
            score_val = mq.get("s", "0")
            # 某些题型可能有逗号分隔的多部分得分
            if isinstance(score_val, str) and "," in score_val:
                try:
                    score_val = sum(float(p) for p in score_val.split(","))
                except ValueError:
                    score_val = 0.0
            else:
                try:
                    score_val = float(score_val)
                except (ValueError, TypeError):
                    score_val = 0.0
            q_scores[qid] = score_val

        # 计算本次答卷的主观题得分 和 其他题型得分
        subj_total = sum(q_scores.get(qid, 0) for qid in subjective_qids)
        other_total = sum(q_scores.get(qid, 0) for qid in other_qids)

        raw_records.append({
            "uid": uid,
            "student_id": student_id,
            "name": name,
            "totalscore": total,
            "q_scores": q_scores,
            "subj_total": subj_total,
            "other_total": other_total,
        })

    # 按学员分组，分别选取主观题最高 和 其他题型最高的答卷合并
    from collections import defaultdict
    by_uid = defaultdict(list)
    for rec in raw_records:
        by_uid[rec["uid"]].append(rec)

    students = []
    for uid, records in by_uid.items():
        if len(records) == 1:
            # 只有一次答卷，直接使用
            best = records[0]
        else:
            # 多次答卷：分别找主观题最高 和 其他题型最高
            best_subj = max(records, key=lambda r: r["subj_total"])
            best_other = max(records, key=lambda r: r["other_total"])

            # 合并两组最佳答卷的 q_scores
            merged_scores = {}
            merged_scores.update(best_subj["q_scores"])   # 先写入主观题最佳
            merged_scores.update(best_other["q_scores"])   # 再用其他题型最佳覆盖/补充

            # 取基础信息（以其他题型最佳的为准，因为它通常决定及格）
            best = {
                "uid": uid,
                "student_id": best_other["student_id"],
                "name": best_other["name"],
                "totalscore": best_subj["subj_total"] + best_other["other_total"],  # 合并后总分
                "q_scores": merged_scores,
            }

            # 日志提示
            if best_subj is not best_other:
                print(f"  学员[{best['name']}] 合并{len(records)}份答卷:"
                      f" 主观题最高={best_subj['subj_total']:.1f}分"
                      f"(来自总分{best_subj['totalscore']:.1f}的答卷)"
                      f" + 其他题最高={best_other['other_total']:.1f}分"
                      f"(来自总分{best_other['totalscore']:.1f}的答卷)")

        students.append({
            "uid": best["uid"],
            "student_id": best["student_id"],
            "name": best["name"],
            "totalscore": best["totalscore"],
            "q_scores": best["q_scores"],
        })

    return students


def build_dataframe(students, questions):
    """
    构建结果 DataFrame
    列：学号/账号、姓名、总分、第1题(题型,分值)、第2题(题型,分值)...
    """
    rows = []
    for s in students:
        row = {
            "学号/账号": s["student_id"],
            "姓名": s["name"],
            "总分": s["totalscore"],
        }
        for q in questions:
            col_name = f"第{q['num']}题({q['type']},{q['points']}分)"
            row[col_name] = s["q_scores"].get(q["id"], 0)
        rows.append(row)

    df = pd.DataFrame(rows)
    df = df.sort_values("总分", ascending=False).reset_index(drop=True)
    return df


def safe_filename(name):
    """将名称转为安全的文件名"""
    return re.sub(r'[\\/:*?"<>|]', '_', name)


def safe_sheet_name(name):
    """将名称转为安全的Sheet名（最长31字符）"""
    safe = re.sub(r'[\\/:*?"<>|]', '_', name)
    return safe[:31]


def process_exam(cid, tid):
    """处理单个考试，返回 (class_name, exam_name, DataFrame)"""
    print(f"\n{'='*60}")
    print(f"正在处理: cid={cid}, tid={tid}")

    # 1. 获取页面信息
    pid, class_name, exam_name = get_page_info(cid, tid)
    if not pid:
        print(f"  [错误] 无法获取pid，请检查Cookie是否有效")
        return None
    print(f"  班级: {class_name}")
    print(f"  考试: {exam_name}")
    print(f"  pid: {pid}")

    # 2. 分页获取数据
    paper_data, examer_data = fetch_all_exam_data(cid, tid, pid)
    print(f"  题目数: {len(paper_data) if paper_data else 0}")
    print(f"  考试人次: {len(examer_data)}")

    if not paper_data or not examer_data:
        print(f"  [警告] 无数据，跳过")
        return None

    # 3. 解析题目结构
    questions = parse_questions(paper_data)
    print(f"  有效题目: {len(questions)}")

    # 4. 解析学员成绩（多答卷按主观题/其他题型分别取最高分合并）
    students = parse_student_scores(examer_data, questions)
    print(f"  独立学员数(去重合并): {len(students)}")

    # 5. 构建 DataFrame
    df = build_dataframe(students, questions)
    print(f"  DataFrame 行数: {len(df)}, 列数: {len(df.columns)}")

    return class_name, exam_name, df


def main():
    """主函数"""
    print("=" * 60)
    print("考试酷成绩批量下载工具")
    print("=" * 60)

    if not EXAMS:
        print("[错误] 未配置考试列表，请在脚本顶部 EXAMS 中添加考试信息")
        return

    # 按班级分组收集结果
    class_data = {}  # class_name -> [(exam_name, df), ...]

    for exam in EXAMS:
        cid = exam["cid"]
        tid = exam["tid"]
        result = process_exam(cid, tid)
        if result:
            class_name, exam_name, df = result
            if class_name not in class_data:
                class_data[class_name] = []
            class_data[class_name].append((exam_name, df))

    # 保存到Excel - 不同班级不同文件，同一班级不同考试不同Sheet
    if not class_data:
        print("\n[错误] 没有获取到任何数据")
        return

    saved_files = []
    for class_name, exams in class_data.items():
        safe_name = safe_filename(class_name)
        filepath = os.path.join(OUTPUT_DIR, f"{safe_name}_成绩.xlsx")

        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            for exam_name, df in exams:
                sheet = safe_sheet_name(exam_name)
                df.to_excel(writer, sheet_name=sheet, index=False)

                # 自动调整列宽
                worksheet = writer.sheets[sheet]
                for col_idx, col in enumerate(df.columns, 1):
                    # 计算最大宽度
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(str(col))
                    )
                    # 中文字符按2个宽度计算
                    col_len = 0
                    for char in str(col):
                        col_len += 2 if ord(char) > 127 else 1
                    for val in df[col].astype(str):
                        val_len = 0
                        for char in val:
                            val_len += 2 if ord(char) > 127 else 1
                        col_len = max(col_len, val_len)
                    adjusted_width = min(col_len + 4, 50)
                    worksheet.column_dimensions[
                        worksheet.cell(row=1, column=col_idx).column_letter
                    ].width = adjusted_width

        saved_files.append(filepath)
        print(f"\n已保存: {filepath}")

    print(f"\n{'='*60}")
    print(f"全部完成！共保存 {len(saved_files)} 个Excel文件:")
    for f in saved_files:
        print(f"  📁 {f}")


if __name__ == "__main__":
    main()


# ===================== 学员成绩汇总处理 =====================

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

# ---- 1. 读取数据 ----
roster = pd.read_excel('C:/Users/Administrator/Desktop/树人计划第六期/成绩处理/学员名单汇总.xlsx', sheet_name='Sheet1')
xl_sren = pd.read_excel('C:/Users/Administrator/Desktop/树人计划第六期/成绩处理/树人计划第6期训练营_成绩.xlsx', sheet_name=None)
xl_toul = pd.read_excel('C:/Users/Administrator/Desktop/树人计划第六期/成绩处理/头狼计划第6期训练营_成绩.xlsx', sheet_name=None)

# ---- 2. 姓名清洗函数：提取末尾的真实姓名 ----
# 例如 "洪梅站-闲宇航" -> "闲宇航"，"广州壹城-中山-沙朗站李汉权" -> "李汉权"
# 规则：如果含 "-"，取最后一个 "-" 之后的部分；否则原样
def clean_name(name):
    name = str(name).strip()
    if '-' in name:
        return name.split('-')[-1].strip()
    return name

# 对成绩文件所有sheet做姓名清洗
def clean_sheet_names(sheets_dict):
    cleaned = {}
    for sheet_name, df in sheets_dict.items():
        df = df.copy()
        if '姓名' in df.columns:
            df['姓名'] = df['姓名'].apply(clean_name)
        cleaned[sheet_name] = df
    return cleaned

xl_sren = clean_sheet_names(xl_sren)
xl_toul = clean_sheet_names(xl_toul)

# ---- 3. 定义课程列表 ----
# 树人课程
sren_courses = list(xl_sren.keys())
# 头狼课程
toul_courses = list(xl_toul.keys())

print('树人课程:', sren_courses)
print('头狼课程:', toul_courses)

# ---- 4. 构建每个课程、每个学员的最高折算分 ----
def calc_exam_full_score(df):
    """
    从列名中计算试卷实际满分。
    列名格式: "第N题(题型,X分)" → 提取每题的 X 并求和
    例如有主观题100分+其他题100分 → 满分200
    """
    import re as _re
    total = 0
    for col in df.columns:
        m = _re.search(r'\((?:[^,]*,)?(\d+(?:\.\d+)?)分\)', str(col))
        if m:
            total += float(m.group(1))
    return total if total > 0 else 100  # 兜底：若无法提取则默认100


def get_max_score_per_person(sheets_dict):
    """
    返回 dict: {sheet_name: {姓名: 折算后最高分}}

    折算规则：
    - 试卷实际满分 = 各题目分值之和（从列名提取）
    - 合并逻辑（已在下载阶段完成）：主观题组与其他题型组分别取最高分后合并
    - 最终折算: 原始合并总分 * 100 / 试卷实际满分
      例如：(非主观最高80 + 主观最高90) * 100 / 200 = 85分
    """
    result = {}
    for sheet_name, df in sheets_dict.items():
        if '姓名' not in df.columns or '总分' not in df.columns:
            continue
        df = df.copy()
        df['总分'] = pd.to_numeric(df['总分'], errors='coerce')

        # 从列名提取试卷实际满分（而非用sheet内最高总分）
        full_score = calc_exam_full_score(df)

        # 按姓名取最高分（下载阶段已按主观/非主观分别取最优合并）
        best = df.groupby('姓名')['总分'].max().reset_index()

        name_score = {}
        for _, row in best.iterrows():
            raw_score = row['总分']
            if pd.isna(raw_score):
                continue
            if full_score > 100:
                # 按实际满分折算为100分制，保留1位小数
                converted = round(raw_score / full_score * 100, 1)
            else:
                converted = raw_score
            name_score[row['姓名']] = converted

        result[sheet_name] = name_score

        print(f'  [{sheet_name}] 试卷满分={full_score}分, 学员数={len(name_score)}')
    return result

sren_scores = get_max_score_per_person(xl_sren)
toul_scores = get_max_score_per_person(xl_toul)

# ---- 5. 判断归属群，区分树人学员 vs 头狼学员 ----
# 树人群: 树人1群
# 头狼群: 头狼1群, 头狼2群
roster['_is_sren'] = roster['归属群'].str.contains('树人', na=False)
roster['_is_toul'] = roster['归属群'].str.contains('头狼', na=False)

# ---- 6. 生成汇总数据 ----
# 根据归属群选用不同的成绩课程
# 树人学员用树人成绩课程，头狼学员用头狼成绩课程
# 但两类学员都在同一名单，需要将对应的列补全

# 将树人课程和头狼课程列名稍作说明（防止重名）
# 检查是否有重名的sheet
sren_set = set(sren_courses)
toul_set = set(toul_courses)
overlap = sren_set & toul_set
print('重名课程:', overlap)

# 重名的课程加前缀区分
def prefix_courses(courses, prefix):
    renamed = {}
    for c in courses:
        if c in overlap:
            renamed[c] = f'[{prefix}]{c}'
        else:
            renamed[c] = c
    return renamed

sren_renamed = prefix_courses(sren_courses, '树人')
toul_renamed = prefix_courses(toul_courses, '头狼')

# ---- 7. 构建输出 DataFrame ----
# 基础列：归属群、序号、姓名、城市、站点
# 然后跟树人各节成绩、头狼各节成绩（每人只显示自己归属的成绩，另一类的成绩留空）

all_cols = ['归属群', '序号', '姓名', '城市', '站点']
# 树人课程列
sren_col_names = [sren_renamed[c] for c in sren_courses]
# 头狼课程列
toul_col_names = [toul_renamed[c] for c in toul_courses]

rows = []
for _, r in roster.iterrows():
    name = r['姓名']
    row = {
        '归属群': r['归属群'],
        '序号': r['序号'],
        '姓名': name,
        '城市': r['城市'],
        '站点': r['站点'],
    }
    
    # 根据归属群填写成绩
    is_sren = r['_is_sren']
    is_toul = r['_is_toul']
    
    # 树人课程成绩
    for c in sren_courses:
        col = sren_renamed[c]
        if is_sren:
            score_map = sren_scores.get(c, {})
            if name in score_map:
                row[col] = score_map[name]
            else:
                row[col] = '未完成'
        else:
            row[col] = ''  # 头狼学员不参加树人课程，留空
    
    # 头狼课程成绩
    for c in toul_courses:
        col = toul_renamed[c]
        if is_toul:
            score_map = toul_scores.get(c, {})
            if name in score_map:
                row[col] = score_map[name]
            else:
                row[col] = '未完成'
        else:
            row[col] = ''  # 树人学员不参加头狼课程，留空
    
    rows.append(row)

df_out = pd.DataFrame(rows)

print('输出列:', df_out.columns.tolist())
print(f'总行数: {len(df_out)}')

# ---- 8. 写出 Excel ----
output_path = 'C:/Users/Administrator/Desktop/树人计划第六期/成绩处理/学员成绩汇总.xlsx'

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '成绩汇总'

# 定义样式
header_font = Font(name='微软雅黑', bold=True, size=10, color='FFFFFF')
header_fill_sren = PatternFill('solid', start_color='2E75B6')   # 蓝色-树人
header_fill_toul = PatternFill('solid', start_color='375623')   # 深绿-头狼
header_fill_base = PatternFill('solid', start_color='404040')   # 深灰-基础信息
cell_font = Font(name='微软雅黑', size=9)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
thin = Side(style='thin', color='AAAAAA')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# 写表头（课程列 + 完成情况备注）
header_row = ['归属群', '序号', '姓名', '城市', '站点'] + sren_col_names + toul_col_names + ['完成情况备注']
for ci, h in enumerate(header_row, 1):
    cell = ws.cell(row=1, column=ci, value=h)
    cell.font = header_font
    cell.alignment = center
    cell.border = border
    if ci <= 5:
        cell.fill = header_fill_base
    elif ci <= 5 + len(sren_col_names):
        cell.fill = header_fill_sren
    elif ci <= 5 + len(sren_col_names) + len(toul_col_names):
        cell.fill = header_fill_toul
    else:
        cell.fill = header_fill_base  # 备注列用深灰表头

# 写数据
sren_member_fill = PatternFill('solid', start_color='DDEEFF')   # 浅蓝-树人成员行
toul_member_fill = PatternFill('solid', start_color='E8F0E0')   # 浅绿-头狼成员行
unfinished_fill = PatternFill('solid', start_color='FFE0E0')    # 浅红-未完成
score_font = Font(name='微软雅黑', size=9)
unfinished_font = Font(name='微软雅黑', size=9, color='CC0000')
low_score_fill = PatternFill('solid', start_color='FFFF00')     # 纯黄色-不及格成绩单元格

# 备注列样式
remark_font_pass = Font(name='微软雅黑', size=9, color='1F7A1F', bold=True)      # 深绿-已通过
remark_font_fail = Font(name='微软雅黑', size=9, color='CC0000', bold=True)      # 红色-有不及格
remark_font_unfinished = Font(name='微软雅黑', size=9, color='7B3F00', bold=True) # 棕色-未完成
remark_fill_pass = PatternFill('solid', start_color='C6EFCE')       # 浅绿-已通过
remark_fill_fail = PatternFill('solid', start_color='FFE0E0')       # 浅红-有不及格
remark_fill_unfinished = PatternFill('solid', start_color='FFF2CC') # 浅黄-有课程未完成

all_col_keys = ['归属群', '序号', '姓名', '城市', '站点'] + sren_col_names + toul_col_names
remark_col_idx = len(all_col_keys) + 1  # 备注列在最后

# 按归属群排序：树人1群，头狼1群，头狼2群
df_out['_sort'] = df_out['归属群'].map({'树人1群': 0, '头狼1群': 1, '头狼2群': 2}).fillna(3)
df_out = df_out.sort_values(['_sort', '序号']).reset_index(drop=True)

for ri, row in df_out.iterrows():
    excel_row = ri + 2
    is_sren = '树人' in str(row.get('归属群', ''))
    row_bg = sren_member_fill if is_sren else toul_member_fill

    # 判断该学员应参加的课程列
    my_course_cols = sren_col_names if is_sren else toul_col_names

    # 写各成绩列，同时记录每列的Excel列索引（供后续标黄不及格用）
    col_idx_map = {}  # col_name -> excel列索引
    for ci, col in enumerate(all_col_keys, 1):
        val = row.get(col, '')
        cell = ws.cell(row=excel_row, column=ci, value=val)
        cell.font = score_font
        cell.alignment = center
        cell.border = border

        # 未完成单元格特殊标色
        if val == '未完成':
            cell.fill = unfinished_fill
            cell.font = unfinished_font
        else:
            cell.fill = row_bg

        col_idx_map[col] = ci

    # ---- 判断完成情况备注逻辑 ----
    # 是否有课程未完成（值为'未完成'）
    has_unfinished = any(row.get(c, '') == '未完成' for c in my_course_cols)
    # 获取所有数值成绩
    numeric_scores = [
        row.get(c, '') for c in my_course_cols
        if isinstance(row.get(c, ''), (int, float)) and not pd.isna(row.get(c, ''))
    ]
    # 是否有低于80分的成绩
    has_low_score = any(s < 80 for s in numeric_scores)

    if has_unfinished:
        # 优先判断：有课程未完成
        remark_val = '有课程未完成'
        remark_cell_font = remark_font_unfinished
        remark_cell_fill = remark_fill_unfinished
    elif has_low_score:
        # 全部完成但有不及格
        remark_val = '有部分成绩不及格'
        remark_cell_font = remark_font_fail
        remark_cell_fill = remark_fill_fail
        # 将不及格（<80分）的成绩单元格标黄色
        for c in my_course_cols:
            v = row.get(c, '')
            if isinstance(v, (int, float)) and not pd.isna(v) and v < 80:
                ci_low = col_idx_map.get(c)
                if ci_low:
                    ws.cell(row=excel_row, column=ci_low).fill = low_score_fill
    else:
        # 全部完成且成绩均 >= 80
        remark_val = '已通过培训全部考核'
        remark_cell_font = remark_font_pass
        remark_cell_fill = remark_fill_pass

    remark_cell = ws.cell(row=excel_row, column=remark_col_idx, value=remark_val)
    remark_cell.font = remark_cell_font
    remark_cell.alignment = center
    remark_cell.border = border
    remark_cell.fill = remark_cell_fill

# 设置列宽
ws.column_dimensions['A'].width = 8
ws.column_dimensions['B'].width = 5
ws.column_dimensions['C'].width = 8
ws.column_dimensions['D'].width = 6
ws.column_dimensions['E'].width = 10
# 课程列宽度
for i in range(6, len(header_row)):  # 不含最后的备注列
    ws.column_dimensions[get_column_letter(i)].width = 14
# 备注列宽度
ws.column_dimensions[get_column_letter(remark_col_idx)].width = 20

# 冻结前5列和首行
ws.freeze_panes = 'F2'

# 第一行行高
ws.row_dimensions[1].height = 40

# ---- 9. 添加第二个sheet：说明 ----
ws2 = wb.create_sheet('折算说明')
ws2['A1'] = '折算规则说明'
ws2['A1'].font = Font(bold=True, size=12)
notes = [
    ['', ''],
    ['规则', '说明'],
    ['多答卷合并', '同一学员有多份答卷时，按题型分组分别取最高分后合并：\n  主观题组(满分100)取主观题得分最高的答卷\n  其他题型组(满分100)取其他题得分最高的答卷'],
    ['折算100分制', '合并后按试卷实际满分折算：\n  折算分 = 合并总分 * 100 / 试卷实际满分（保留1位小数）\n  例：(非主观最高80 + 主观最高90) × 100 / 200 = 85分'],
    ['试卷实际满分', '由各题目分值之和确定（从列名提取），含主观题的试卷通常为200分'],
    ['未完成', '学员在名单中但该节课无成绩记录，显示"未完成"'],
    ['空白', '该节课不属于该学员所在群组，不做统计'],
    ['', ''],
    ['完成情况备注说明', ''],
    ['已通过培训全部考核', '全部课程均已完成，且所有成绩 >= 80分'],
    ['有部分成绩不及格', '全部课程已完成，但存在成绩 < 80分（不及格单元格标黄）'],
    ['有课程未完成', '存在未完成的课程（成绩记录缺失）'],
    ['', ''],
    ['树人成绩说明', ''],
]
for c in sren_courses:
    df = xl_sren[c] if c in xl_sren else None
    if df is not None and '总分' in df.columns:
        full_score = calc_exam_full_score(df)
        max_raw = pd.to_numeric(df['总分'], errors='coerce').max()
        notes.append([c, f'试卷满分={full_score}分, 合并后最高原始分={max_raw}分' + ('（需折算）' if full_score > 100 else '（无需折算）')])
notes.append(['', ''])
notes.append(['头狼成绩说明', ''])
for c in toul_courses:
    df = xl_toul[c] if c in xl_toul else None
    if df is not None and '总分' in df.columns:
        full_score = calc_exam_full_score(df)
        max_raw = pd.to_numeric(df['总分'], errors='coerce').max()
        col_label = toul_renamed[c]
        notes.append([col_label, f'试卷满分={full_score}分, 合并后最高原始分={max_raw}分' + ('（需折算）' if full_score > 100 else '（无需折算）')])

for r_data in notes:
    ws2.append(r_data)

ws2.column_dimensions['A'].width = 30
ws2.column_dimensions['B'].width = 40

wb.save(output_path)
print(f'\n已保存至: {output_path}')

# 打印匹配统计
print('\n=== 匹配统计 ===')
for c in sren_courses:
    score_map = sren_scores.get(c, {})
    matched = sum(1 for _, r in roster[roster['_is_sren']].iterrows() if r['姓名'] in score_map)
    total_sren = roster['_is_sren'].sum()
    print(f'[树人]{c}: {matched}/{total_sren} 人有成绩')

for c in toul_courses:
    score_map = toul_scores.get(c, {})
    matched = sum(1 for _, r in roster[roster['_is_toul']].iterrows() if r['姓名'] in score_map)
    total_toul = roster['_is_toul'].sum()
    print(f'[头狼]{c}: {matched}/{total_toul} 人有成绩')
