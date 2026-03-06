#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教师科研统计分析工具
根据教师姓名，统计文章数量、项目经费等信息，并生成报告和散点图。
"""

import pandas as pd
# import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import re
import os
import sys
import platform
import shutil
import hashlib
import subprocess
from pathlib import Path
from datetime import datetime
from pypinyin import lazy_pinyin

# ============================================================
# 配置区：在此修改要查询的教师姓名
# ============================================================
TEACHER_NAMES = ["高鹤","朱宗宏", "何林"]  # 修改此列表，可以是一个或多个教师
# TEACHER_NAMES = ["高鹤"]
# ============================================================

# 设置matplotlib中文字体（跨平台自动检测可用中文字体）
def _setup_chinese_font():
    import matplotlib.font_manager as fm
    # 按优先级列出各平台常见中文字体
    candidates = [
        'SimHei',               # Windows
        'Microsoft YaHei',      # Windows
        'PingFang SC',          # macOS
        'Heiti SC',             # macOS
        'STHeiti',              # macOS
        'Arial Unicode MS',    # macOS
        'WenQuanYi Micro Hei', # Linux
        'Noto Sans CJK SC',    # Linux
        'Droid Sans Fallback', # Linux
    ]
    available = {f.name for f in fm.fontManager.ttflist}
    for font in candidates:
        if font in available:
            plt.rcParams['font.sans-serif'] = [font, 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            return
    # 未找到已知中文字体，使用默认列表
    plt.rcParams['font.sans-serif'] = candidates + ['DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False

_setup_chinese_font()

current_year = datetime.now().year
RECENT_YEARS = 5  # 近五年

# 数据路径（兼容 PyInstaller 打包：冻结时从 sys._MEIPASS 读取 bundled 资源）
def _get_base_path():
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent

DATA_PATH = _get_base_path() / "data"
OUTPUT_PATH = Path(__file__).parent / "output"


# ============================================================
# 数据文件哈希与缓存
# ============================================================
DATA_FILES = [
    "在编信息汇总.xlsx",
    "人才信息汇总.xlsx",
    "成果批量导出.xlsx",
    "纵向项目.xls",
    "横向项目.xls",
]

HASH_PREFIX = "# data_hash:"


def compute_data_hash(file_paths=None):
    """计算data目录下所有关键数据文件的综合哈希值（含 RECENT_YEARS 和 current_year）"""
    h = hashlib.md5()
    # 将近几年参数纳入哈希，切换时自动使缓存失效
    h.update(f"current_year={current_year},RECENT_YEARS={RECENT_YEARS}".encode('utf-8'))
    if file_paths is None:
        paths = [(fname, DATA_PATH / fname) for fname in sorted(DATA_FILES)]
    else:
        paths = [(Path(p).name, Path(p)) for p in sorted(file_paths.values(), key=str)]
    for fname, fpath in paths:
        if fpath.exists():
            h.update(fname.encode('utf-8'))
            with open(fpath, 'rb') as f:
                for chunk in iter(lambda: f.read(8192), b''):
                    h.update(chunk)
    return h.hexdigest()


def save_csv_with_hash(df, csv_path, data_hash):
    """保存CSV文件，第一行写入数据哈希值"""
    os.makedirs(csv_path.parent, exist_ok=True)
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(f"{HASH_PREFIX}{data_hash}\n")
        df.to_csv(f, index=False)


def read_hash_from_csv(csv_path):
    """从CSV文件第一行读取数据哈希值，若不存在则返回None"""
    if not csv_path.exists():
        return None
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        first_line = f.readline().strip()
    if first_line.startswith(HASH_PREFIX):
        return first_line[len(HASH_PREFIX):]
    return None


def load_csv_with_hash(csv_path):
    """加载含哈希行的CSV文件，跳过第一行哈希注释"""
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        first_line = f.readline().strip()
        if first_line.startswith(HASH_PREFIX):
            return pd.read_csv(f)
        else:
            # 没有哈希行，回到开头正常读取
            f.seek(0)
            return pd.read_csv(f)


# ============================================================
# 数据加载
# ============================================================
def load_excel(file_path):
    """加载Excel文件，读取所有sheet并合并"""
    engine = None
    if file_path.suffix == '.xlsx':
        engine = "openpyxl"
    elif file_path.suffix == '.xls':
        engine = "xlrd"
    else:
        return None
    sheets = pd.read_excel(file_path, engine=engine, dtype=str, na_values=["NULL"], sheet_name=None)
    return pd.concat(sheets.values(), ignore_index=True)


def load_all_data(file_paths=None):
    """加载所有数据文件"""
    print("加载数据文件...")
    if file_paths is None:
        file_paths = {
            "在编信息汇总": DATA_PATH / "在编信息汇总.xlsx",
            "人才信息汇总": DATA_PATH / "人才信息汇总.xlsx",
            "成果批量导出": DATA_PATH / "成果批量导出.xlsx",
            "纵向项目": DATA_PATH / "纵向项目.xls",
            "横向项目": DATA_PATH / "横向项目.xls",
        }
    people_df = load_excel(Path(file_paths["在编信息汇总"]))
    talent_df = load_excel(Path(file_paths["人才信息汇总"]))
    papers_df = load_excel(Path(file_paths["成果批量导出"]))
    vertical_df = load_excel(Path(file_paths["纵向项目"]))
    horizontal_df = load_excel(Path(file_paths["横向项目"]))
    print("数据加载完成。")
    return people_df, talent_df, papers_df, vertical_df, horizontal_df


# ============================================================
# 教师筛选
# ============================================================
def filter_teachers(people_df):
    """筛选岗位子类别包含'教学科研'或'工程实验'的教师"""
    mask = people_df['岗位子类别'].str.contains('教学科研|工程实验', na=False)
    filtered = people_df[mask].copy()
    print(f"筛选教学科研/工程实验教师：{len(filtered)} 人")
    return filtered


def validate_teacher_names(teacher_names, file_paths=None):
    """检查教师姓名是否在在编信息中，返回不在名单中的教师列表"""
    if file_paths is None:
        people_path = DATA_PATH / "在编信息汇总.xlsx"
    else:
        people_path = Path(file_paths["在编信息汇总"])
    people_df = load_excel(people_path)
    all_names = set(people_df['姓名'].dropna().astype(str))
    not_found = [n for n in teacher_names if n not in all_names]
    return not_found


# ============================================================
# 拼音生成
# ============================================================
def generate_pinyin_formats(chinese_name):
    """为中文姓名生成36种拼音英文格式"""
    if pd.isna(chinese_name) or chinese_name == '':
        return []

    name = str(chinese_name).strip()
    pinyin_list = lazy_pinyin(name)
    if not pinyin_list or len(pinyin_list) < 2:
        if len(pinyin_list) == 1:
            return [pinyin_list[0].capitalize()]
        return []

    surname = pinyin_list[0].capitalize()
    given_name_parts = [part.capitalize() for part in pinyin_list[1:]]
    given_name_combined = ''.join(given_name_parts)
    given_name_hyphen = '-'.join(given_name_parts)

    given_initial = given_name_parts[0][0].upper()
    surname_initial = surname[0].upper()
    all_initials = [part[0].upper() for part in given_name_parts]
    all_initials_spaced = '. '.join(all_initials) + '.'
    all_initials_hyphen = '.-'.join(all_initials) + '.'

    formats = [
        f"{surname} {given_name_combined}",
        f"{given_name_combined} {surname}",
        f"{surname} {given_name_hyphen}",
        f"{given_name_hyphen} {surname}",
        f"{given_initial}. {surname}",
        f"{surname} {given_initial}.",
        f"{given_name_combined} {surname_initial}.",
        f"{given_name_hyphen} {surname_initial}.",
        f"{surname} {all_initials_spaced}",
        f"{all_initials_spaced} {surname}",
        f"{surname} {all_initials_hyphen}",
        f"{all_initials_hyphen} {surname}",
        f"{surname}, {given_name_combined}",
        f"{given_name_combined}, {surname}",
        f"{surname}, {given_name_hyphen}",
        f"{given_name_hyphen}, {surname}",
        f"{given_initial}., {surname}",
        f"{surname}, {given_initial}.",
        f"{given_name_combined}, {surname_initial}.",
        f"{given_name_hyphen}, {surname_initial}.",
        f"{surname}, {all_initials_spaced}",
        f"{all_initials_spaced}, {surname}",
        f"{surname}, {all_initials_hyphen}",
        f"{all_initials_hyphen}, {surname}",
        f"{surname},{given_name_combined}",
        f"{given_name_combined},{surname}",
        f"{surname},{given_name_hyphen}",
        f"{given_name_hyphen},{surname}",
        f"{given_initial}.,{surname}",
        f"{surname},{given_initial}.",
        f"{given_name_combined},{surname_initial}.",
        f"{given_name_hyphen},{surname_initial}.",
        f"{surname},{all_initials_spaced}",
        f"{all_initials_spaced},{surname}",
        f"{surname},{all_initials_hyphen}",
        f"{all_initials_hyphen},{surname}",
    ]
    return formats


# ============================================================
# 文章去重
# ============================================================
def deduplicate_papers(papers_df):
    """智能去重：按题名分组，保留最佳记录"""
    title_col = '题名'

    def smart_dedup(group):
        if len(group) == 1:
            return group.iloc[0]
        scores = []
        for idx, row in group.iterrows():
            score = 0
            if pd.notna(row.get('成果归属学者', '')) and str(row.get('成果归属学者', '')) not in ('nan', ''):
                score += 3
            if pd.notna(row.get('作者', '')) and str(row.get('作者', '')) not in ('nan', ''):
                score += 2
            if pd.notna(row.get('出版日期', '')) and str(row.get('出版日期', '')) not in ('nan', ''):
                score += 1
            score += 0.1 / (idx + 1)
            scores.append((score, idx))
        best_idx = max(scores, key=lambda x: x[0])[1]
        return group.loc[best_idx]

    cleaned = papers_df.groupby(title_col, group_keys=False).apply(smart_dedup, include_groups=False).reset_index(drop=True)
    print(f"文章去重: {len(papers_df)} -> {len(cleaned)}")
    return cleaned


# ============================================================
# 文章匹配统计
# ============================================================
def extract_year(year_str):
    """从年份字段提取整数年"""
    if pd.isna(year_str) or str(year_str).strip() == '' or str(year_str) == 'nan':
        return None
    try:
        return int(float(str(year_str).strip()))
    except (ValueError, TypeError):
        m = re.match(r'(\d{4})', str(year_str).strip())
        if m:
            return int(m.group(1))
        return None


def match_papers_for_teachers(teachers_df, papers_df):
    """
    为所有老师匹配文章，返回统计结果DataFrame。
    列：姓名, 总文章数量, 第一署名单位文章数量, 通讯作者文章数量, 近五年文章数量, 近五年第一署名单位文章数量, 近五年通讯作者文章数量
    """
    print("开始匹配文章...")

    # 预生成所有教师的拼音
    teacher_pinyin = {}
    for _, row in teachers_df.iterrows():
        name = row['姓名']
        if pd.notna(name):
            teacher_pinyin[name] = generate_pinyin_formats(name)

    # 提取文章年份
    papers_df = papers_df.copy()
    papers_df['_year'] = papers_df['年'].apply(extract_year)

    # 预处理作者和成果归属学者字段
    author_col_vals = papers_df['作者'].fillna('').str.lower().tolist()
    scholar_col_vals = papers_df['成果归属学者'].fillna('').tolist()
    order_col_vals = papers_df['本机构署名顺序'].fillna('').tolist()
    corresponding_col_vals = papers_df['通讯作者'].fillna('').str.lower().tolist()
    corresponding_unit_col_vals = papers_df['通讯作者单位'].fillna('').str.lower().tolist()
    year_vals = papers_df['_year'].tolist()

    five_year_start = current_year - RECENT_YEARS + 1

    results = []
    total = len(teacher_pinyin)
    for i, (name, formats) in enumerate(teacher_pinyin.items()):
        if (i + 1) % 50 == 0 or i == 0:
            print(f"  匹配进度: {i+1}/{total}")

        total_papers = 0
        first_order_papers = 0
        corresponding_papers = 0
        recent_papers = 0
        recent_first_order_papers = 0
        recent_corresponding_papers = 0

        if not formats:
            results.append({
                '姓名': name,
                '总文章数量': 0,
                '第一署名单位文章数量': 0,
                '通讯作者文章数量': 0,
                '近五年文章数量': 0,
                '近五年第一署名单位文章数量': 0,
                '近五年通讯作者文章数量': 0,
            })
            continue

        formats_lower = [f.lower() for f in formats if f]

        for j in range(len(papers_df)):
            author_str = author_col_vals[j]
            scholar_str = scholar_col_vals[j]

            # 检查是否匹配（在作者或成果归属学者中查找拼音格式，或中文名直接在成果归属学者中）
            matched = False
            if name in scholar_str:
                matched = True
            else:
                for fmt in formats_lower:
                    if fmt in author_str:
                        matched = True
                        break
                    if fmt in scholar_str.lower():
                        matched = True
                        break

            if matched:
                total_papers += 1
                is_first = (order_col_vals[j] == '第一署名顺序')
                year = year_vals[j]

                # 判断是否为通讯作者（匹配通讯作者列或通讯作者单位列）
                corr_str = corresponding_col_vals[j]
                corr_unit_str = corresponding_unit_col_vals[j]
                is_corresponding = False
                if name in corr_str or name in corr_unit_str:
                    is_corresponding = True
                else:
                    for fmt in formats_lower:
                        fmt_no_dot = fmt.replace('.', '')
                        if fmt in corr_str or fmt in corr_unit_str \
                                or fmt_no_dot in corr_str or fmt_no_dot in corr_unit_str:
                            is_corresponding = True
                            break

                if is_first:
                    first_order_papers += 1
                if is_corresponding:
                    corresponding_papers += 1

                if year is not None and year >= five_year_start:
                    recent_papers += 1
                    if is_first:
                        recent_first_order_papers += 1
                    if is_corresponding:
                        recent_corresponding_papers += 1

        results.append({
            '姓名': name,
            '总文章数量': total_papers,
            '第一署名单位文章数量': first_order_papers,
            '通讯作者文章数量': corresponding_papers,
            '近五年文章数量': recent_papers,
            '近五年第一署名单位文章数量': recent_first_order_papers,
            '近五年通讯作者文章数量': recent_corresponding_papers,
        })

    print(f"文章匹配完成，共{len(results)}位教师。")
    return pd.DataFrame(results)


# ============================================================
# 经费统计
# ============================================================
def clean_funding(s):
    """清理经费金额字段，转为浮点数（单位：万元）"""
    if pd.isna(s) or str(s).strip() in ('', 'nan'):
        return 0.0
    s = re.sub(r'[万元￥¥,，\s]', '', str(s).strip())
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def extract_date_year(date_str):
    """从日期字符串中提取年份"""
    if pd.isna(date_str) or str(date_str).strip() in ('', 'nan'):
        return None
    m = re.match(r'(\d{4})', str(date_str).strip())
    return int(m.group(1)) if m else None


def compute_funding_stats(teachers_df, vertical_df, horizontal_df):
    """
    计算每位教师的经费统计。
    返回DataFrame: 姓名, NSFC生涯数量, NSFC近五年数量, 生涯总经费, 近五年总经费
    """
    print("开始统计项目经费...")
    teacher_names_set = set(teachers_df['姓名'].dropna().astype(str))
    five_year_start = current_year - RECENT_YEARS + 1

    # --- 纵向项目 ---
    v_all = vertical_df.copy()
    v_all['_year'] = v_all['立项日期'].apply(extract_date_year)
    v_all['_funding'] = v_all['批准经费'].apply(clean_funding)
    v_teacher = v_all[v_all['负责人'].isin(teacher_names_set)].copy()

    # NSFC项目（不考虑状态）
    v_nsfc = v_teacher[v_teacher['项目分类'].str.startswith('国家自然科学基金', na=False)].copy()

    # --- 横向项目 ---
    h_all = horizontal_df.copy()
    h_all['_year'] = h_all['立项日期'].apply(extract_date_year)
    h_all['_funding'] = h_all['批准经费'].apply(clean_funding)
    h_teacher = h_all[h_all['负责人'].isin(teacher_names_set)].copy()

    results = []
    for name in teacher_names_set:
        # NSFC统计
        nsfc_mine = v_nsfc[v_nsfc['负责人'] == name]
        nsfc_total = len(nsfc_mine)
        nsfc_recent = len(nsfc_mine[nsfc_mine['_year'].apply(lambda y: y is not None and y >= five_year_start)])

        # 生涯总经费 = 纵向 + 横向
        v_mine = v_teacher[v_teacher['负责人'] == name]
        h_mine = h_teacher[h_teacher['负责人'] == name]
        total_funding = v_mine['_funding'].sum() + h_mine['_funding'].sum()

        # 近五年总经费
        v_recent_funding = 0.0
        h_recent_funding = 0.0
        if len(v_mine) > 0:
            v_recent = v_mine[v_mine['_year'].apply(lambda y: y is not None and y >= five_year_start)]
            v_recent_funding = v_recent['_funding'].sum() if len(v_recent) > 0 else 0.0
        if len(h_mine) > 0:
            h_recent = h_mine[h_mine['_year'].apply(lambda y: y is not None and y >= five_year_start)]
            h_recent_funding = h_recent['_funding'].sum() if len(h_recent) > 0 else 0.0
        recent_funding = v_recent_funding + h_recent_funding

        results.append({
            '姓名': name,
            'NSFC生涯数量': nsfc_total,
            'NSFC近五年数量': nsfc_recent,
            '生涯总经费(万元)': round(total_funding, 2),
            '近五年总经费(万元)': round(recent_funding, 2),
        })

    print(f"经费统计完成，共{len(results)}位教师。")
    return pd.DataFrame(results)


# ============================================================
# 人才信息与个人信息查询
# ============================================================
def get_teacher_info(name, people_df, talent_df):
    """获取教师的人才称号、年龄、毕业学校、学位"""
    info = {}

    # 在编信息
    row = people_df[people_df['姓名'] == name]
    if len(row) > 0:
        row = row.iloc[0]
        # 计算年龄
        birth = row.get('出生日期', '')
        age = ''
        if pd.notna(birth) and str(birth) not in ('', 'nan'):
            try:
                birth_date = pd.to_datetime(str(birth))
                age = current_year - birth_date.year
            except Exception:
                age = ''
        info['年龄'] = age
        info['毕业学校'] = row.get('毕业学校', '')
        info['学位'] = row.get('学位', '')
        info['岗位子类别'] = row.get('岗位子类别', '')
        info['专业技术职务'] = row.get('专业技术职务', '')
        info['一级学科'] = row.get('一级学科', '')
    else:
        info['年龄'] = ''
        info['毕业学校'] = ''
        info['学位'] = ''
        info['岗位子类别'] = ''
        info['专业技术职务'] = ''
        info['一级学科'] = ''

    # 人才信息
    talent_rows = talent_df[talent_df['姓名'].str.contains(name, na=False)]
    if len(talent_rows) > 0:
        honors = []
        for _, tr in talent_rows.iterrows():
            h = tr.get('人才/荣誉称号', '')
            if pd.notna(h) and str(h) not in ('', 'nan'):
                honors.append(str(h))
        info['人才称号'] = '；'.join(honors) if honors else '无'
        info['是否人才'] = '是' if honors else '否'
    else:
        info['人才称号'] = '无'
        info['是否人才'] = '否'

    return info


# ============================================================
# 散点图绘制
# ============================================================
SCATTER_CONFIGS = [
    ('生涯总经费(万元)', '总文章数量'),
    ('生涯总经费(万元)', '第一署名单位文章数量'),
    ('生涯总经费(万元)', '通讯作者文章数量'),
    ('生涯总经费(万元)', '近五年文章数量'),
    ('生涯总经费(万元)', '近五年第一署名单位文章数量'),
    ('生涯总经费(万元)', '近五年通讯作者文章数量'),
    ('近五年总经费(万元)', '总文章数量'),
    ('近五年总经费(万元)', '第一署名单位文章数量'),
    ('近五年总经费(万元)', '通讯作者文章数量'),
    ('近五年总经费(万元)', '近五年文章数量'),
    ('近五年总经费(万元)', '近五年第一署名单位文章数量'),
    ('近五年总经费(万元)', '近五年通讯作者文章数量'),
]


def draw_scatter_single(all_stats, teacher_name, output_dir):
    """为单个教师绘制8张散点图（灰色背景+彩色突出）"""
    os.makedirs(output_dir, exist_ok=True)

    for x_col, y_col in SCATTER_CONFIGS:
        fig, ax = plt.subplots(figsize=(10, 7))

        # 绘制所有教师为灰色
        ax.scatter(all_stats[x_col], all_stats[y_col],
                   c='lightgray', s=30, alpha=0.6, edgecolors='gray', linewidths=0.5,
                   label='所有教师', zorder=1)

        # 突出目标教师
        target = all_stats[all_stats['姓名'] == teacher_name]
        if len(target) > 0:
            tx = target[x_col].values[0]
            ty = target[y_col].values[0]
            ax.scatter(tx, ty, c='red', s=120, edgecolors='darkred', linewidths=1.5,
                       label=teacher_name, zorder=3)
            ax.annotate(f'{teacher_name}\n({tx:.1f}, {ty})',
                        xy=(tx, ty), xytext=(15, 15),
                        textcoords='offset points', fontsize=9,
                        bbox=dict(boxstyle='round,pad=0.3', facecolor='yellow', alpha=0.8),
                        arrowprops=dict(arrowstyle='->', color='red'),
                        zorder=4)

        ax.set_xlabel(x_col, fontsize=12)
        ax.set_ylabel(y_col, fontsize=12)
        ax.set_title(f'{y_col} vs {x_col}', fontsize=14)
        ax.legend(loc='upper left', fontsize=10)
        ax.grid(True, alpha=0.3)
        plt.tight_layout()

        fname = f"{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
        fig.savefig(os.path.join(output_dir, fname), dpi=150, bbox_inches='tight')
        plt.close(fig)


def draw_scatter_comparison(all_stats, teacher_names, output_dir):
    """为多个教师绘制对比散点图（灰色背景+多彩色突出）"""
    os.makedirs(output_dir, exist_ok=True)

    colors = ['red', 'blue', 'green', 'orange', 'purple', 'brown', 'deeppink', 'teal']

    for x_col, y_col in SCATTER_CONFIGS:
        fig, ax = plt.subplots(figsize=(10, 7))

        # 灰色背景
        ax.scatter(all_stats[x_col], all_stats[y_col],
                   c='lightgray', s=30, alpha=0.6, edgecolors='gray', linewidths=0.5,
                   label='所有教师', zorder=1)

        # 多个教师彩色突出
        for idx, name in enumerate(teacher_names):
            color = colors[idx % len(colors)]
            target = all_stats[all_stats['姓名'] == name]
            if len(target) > 0:
                tx = target[x_col].values[0]
                ty = target[y_col].values[0]
                ax.scatter(tx, ty, c=color, s=120, edgecolors='black', linewidths=1.5,
                           label=name, zorder=3)
                ax.annotate(f'{name}\n({tx:.1f}, {ty})',
                            xy=(tx, ty), xytext=(15, 15 + idx * 25),
                            textcoords='offset points', fontsize=9,
                            bbox=dict(boxstyle='round,pad=0.3', facecolor=color, alpha=0.3),
                            arrowprops=dict(arrowstyle='->', color=color),
                            zorder=4)

        ax.set_xlabel(x_col, fontsize=12)
        ax.set_ylabel(y_col, fontsize=12)
        ax.set_title(f'{y_col} vs {x_col}（对比）', fontsize=14)
        ax.legend(loc='upper left', fontsize=10)
        ax.grid(True, alpha=0.3)
        plt.tight_layout()

        fname = f"对比_{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
        fig.savefig(os.path.join(output_dir, fname), dpi=150, bbox_inches='tight')
        plt.close(fig)


# ============================================================
# Markdown转PDF
# ============================================================
def _get_pandoc_path():
    """查找 pandoc 可执行文件（优先使用打包的版本）"""
    app_dir = str(_get_base_path())
    exe_name = "pandoc.exe" if platform.system() == "Windows" else "pandoc"
    bundled = os.path.join(app_dir, exe_name)
    if os.path.isfile(bundled):
        return bundled
    return shutil.which("pandoc")


def md_to_docx(md_path):
    """使用pandoc将Markdown转为DOCX"""
    pandoc = _get_pandoc_path()
    if pandoc is None:
        print("  警告：未找到pandoc，跳过DOCX生成。")
        return
    md_abs = os.path.abspath(md_path)
    md_dir = os.path.dirname(md_abs)
    md_name = os.path.basename(md_abs)
    docx_name = md_name.replace('.md', '.docx')
    try:
        subprocess.run([
            pandoc, md_name, '-o', docx_name,
        ], check=True, capture_output=True, text=True, cwd=md_dir)
        print(f"  已生成DOCX: {os.path.join(md_dir, docx_name)}")
    except subprocess.CalledProcessError as e:
        print(f"  DOCX生成失败: {e.stderr[:200] if e.stderr else e}")
    except FileNotFoundError:
        print("  警告：未找到pandoc，跳过DOCX生成。")


# ============================================================
# 报告生成
# ============================================================
def generate_individual_report(name, paper_stats, funding_stats, teacher_info, output_dir):
    """生成单个教师的Markdown报告"""
    os.makedirs(output_dir, exist_ok=True)

    p = paper_stats[paper_stats['姓名'] == name]
    f = funding_stats[funding_stats['姓名'] == name]

    lines = []
    lines.append(f"# {name} 科研统计报告\n")
    lines.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    lines.append("## 基本信息\n")
    lines.append(f"| 项目 | 内容 |")
    lines.append(f"|------|------|")
    lines.append(f"| 姓名 | {name} |")
    lines.append(f"| 年龄 | {teacher_info.get('年龄', '')} |")
    lines.append(f"| 毕业学校 | {teacher_info.get('毕业学校', '')} |")
    lines.append(f"| 学位 | {teacher_info.get('学位', '')} |")
    lines.append(f"| 岗位子类别 | {teacher_info.get('岗位子类别', '')} |")
    lines.append(f"| 专业技术职务 | {teacher_info.get('专业技术职务', '')} |")
    lines.append(f"| 一级学科 | {teacher_info.get('一级学科', '')} |")
    lines.append(f"| 是否人才 | {teacher_info.get('是否人才', '否')} |")
    lines.append(f"| 人才称号 | {teacher_info.get('人才称号', '无')} |")
    lines.append("")

    lines.append("## 文章统计\n")
    lines.append(f"| 指标 | 数量 |")
    lines.append(f"|------|------|")
    if len(p) > 0:
        pr = p.iloc[0]
        lines.append(f"| 总文章数量 | {int(pr['总文章数量'])} |")
        lines.append(f"| 第一署名单位文章数量 | {int(pr['第一署名单位文章数量'])} |")
        lines.append(f"| 通讯作者文章数量 | {int(pr['通讯作者文章数量'])} |")
        lines.append(f"| 近五年文章数量 | {int(pr['近五年文章数量'])} |")
        lines.append(f"| 近五年第一署名单位文章数量 | {int(pr['近五年第一署名单位文章数量'])} |")
        lines.append(f"| 近五年通讯作者文章数量 | {int(pr['近五年通讯作者文章数量'])} |")
    else:
        lines.append(f"| 总文章数量 | 0 |")
        lines.append(f"| 第一署名单位文章数量 | 0 |")
        lines.append(f"| 通讯作者文章数量 | 0 |")
        lines.append(f"| 近五年文章数量 | 0 |")
        lines.append(f"| 近五年第一署名单位文章数量 | 0 |")
        lines.append(f"| 近五年通讯作者文章数量 | 0 |")
    lines.append("")

    lines.append("## 项目经费统计\n")
    lines.append(f"| 指标 | 数值 |")
    lines.append(f"|------|------|")
    if len(f) > 0:
        fr = f.iloc[0]
        lines.append(f"| NSFC生涯项目数量 | {int(fr['NSFC生涯数量'])} |")
        lines.append(f"| NSFC近五年项目数量 | {int(fr['NSFC近五年数量'])} |")
        lines.append(f"| 生涯总经费（万元） | {fr['生涯总经费(万元)']:.2f} |")
        lines.append(f"| 近五年总经费（万元） | {fr['近五年总经费(万元)']:.2f} |")
    else:
        lines.append(f"| NSFC生涯项目数量 | 0 |")
        lines.append(f"| NSFC近五年项目数量 | 0 |")
        lines.append(f"| 生涯总经费（万元） | 0.00 |")
        lines.append(f"| 近五年总经费（万元） | 0.00 |")
    lines.append("")

    # 添加散点图
    lines.append("## 散点图\n")
    for x_col, y_col in SCATTER_CONFIGS:
        fname = f"{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
        img_path = os.path.join(output_dir, fname)
        if os.path.exists(img_path):
            lines.append(f"### {y_col} vs {x_col}\n")
            lines.append(f"![{y_col} vs {x_col}]({fname})\n")

    report_path = os.path.join(output_dir, f"{name}_统计报告.md")
    with open(report_path, 'w', encoding='utf-8') as fout:
        fout.write('\n'.join(lines))
    print(f"  已生成报告: {report_path}")
    md_to_docx(report_path)


def generate_comparison_report(names, paper_stats, funding_stats, teacher_infos, output_dir):
    """生成多教师对比的Markdown报告"""
    os.makedirs(output_dir, exist_ok=True)

    lines = []
    lines.append(f"# 教师科研统计对比报告\n")
    lines.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    lines.append(f"对比教师：{'、'.join(names)}\n")

    # 表头
    header = "| 指标 | " + " | ".join(names) + " |"
    sep = "|------|" + "|".join(["------"] * len(names)) + "|"

    lines.append("## 基本信息对比\n")
    lines.append(header)
    lines.append(sep)

    info_fields = [
        ('年龄', '年龄'), ('毕业学校', '毕业学校'), ('学位', '学位'),
        ('岗位子类别', '岗位子类别'), ('专业技术职务', '专业技术职务'),
        ('一级学科', '一级学科'), ('是否人才', '是否人才'), ('人才称号', '人才称号'),
    ]
    for label, key in info_fields:
        vals = [str(teacher_infos[n].get(key, '')) for n in names]
        lines.append(f"| {label} | " + " | ".join(vals) + " |")
    lines.append("")

    lines.append("## 文章统计对比\n")
    lines.append(header)
    lines.append(sep)
    paper_fields = ['总文章数量', '第一署名单位文章数量', '通讯作者文章数量', '近五年文章数量', '近五年第一署名单位文章数量', '近五年通讯作者文章数量']
    for field in paper_fields:
        vals = []
        for n in names:
            row = paper_stats[paper_stats['姓名'] == n]
            if len(row) > 0:
                vals.append(str(int(row.iloc[0][field])))
            else:
                vals.append('0')
        lines.append(f"| {field} | " + " | ".join(vals) + " |")
    lines.append("")

    lines.append("## 项目经费统计对比\n")
    lines.append(header)
    lines.append(sep)
    funding_fields = [
        ('NSFC生涯数量', 'NSFC生涯项目数量', True),
        ('NSFC近五年数量', 'NSFC近五年项目数量', True),
        ('生涯总经费(万元)', '生涯总经费（万元）', False),
        ('近五年总经费(万元)', '近五年总经费（万元）', False),
    ]
    for col, label, is_int in funding_fields:
        vals = []
        for n in names:
            row = funding_stats[funding_stats['姓名'] == n]
            if len(row) > 0:
                v = row.iloc[0][col]
                vals.append(str(int(v)) if is_int else f"{v:.2f}")
            else:
                vals.append('0' if is_int else '0.00')
        lines.append(f"| {label} | " + " | ".join(vals) + " |")
    lines.append("")

    lines.append("")

    # 添加对比散点图
    lines.append("## 对比散点图\n")
    for x_col, y_col in SCATTER_CONFIGS:
        fname = f"对比_{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
        img_path = os.path.join(output_dir, fname)
        if os.path.exists(img_path):
            lines.append(f"### {y_col} vs {x_col}\n")
            lines.append(f"![{y_col} vs {x_col}]({fname})\n")

    report_path = os.path.join(output_dir, "对比报告.md")
    with open(report_path, 'w', encoding='utf-8') as fout:
        fout.write('\n'.join(lines))
    print(f"  已生成对比报告: {report_path}")
    md_to_docx(report_path)


# ============================================================
# 主程序
# ============================================================
def main(teacher_names=None, file_paths=None, output_path=None):
    if teacher_names is None:
        teacher_names = TEACHER_NAMES
    if output_path is not None:
        output_path = Path(output_path)
    else:
        output_path = OUTPUT_PATH
    if not teacher_names:
        print("错误：请在脚本顶部设置 TEACHER_NAMES 列表。")
        sys.exit(1)

    print(f"查询教师：{'、'.join(teacher_names)}")
    print(f"当前年份：{current_year}，近五年范围：{current_year - RECENT_YEARS + 1}-{current_year}")
    print("=" * 60)

    # 1. 计算数据文件哈希值
    print("计算数据文件哈希值...")
    data_hash = compute_data_hash(file_paths)
    print(f"数据哈希: {data_hash}")

    # 创建 output 及 cache 子目录
    cache_path = output_path / "cache"
    os.makedirs(output_path, exist_ok=True)
    os.makedirs(cache_path, exist_ok=True)

    paper_csv = cache_path / "文章统计.csv"
    funding_csv = cache_path / "经费统计.csv"

    # 检查缓存：如果输出文件存在且哈希一致，直接读取
    cached_paper_hash = read_hash_from_csv(paper_csv)
    cached_funding_hash = read_hash_from_csv(funding_csv)

    EXPECTED_PAPER_COLS = {'第一署名单位文章数量', '近五年第一署名单位文章数量', '通讯作者文章数量', '近五年通讯作者文章数量'}

    if cached_paper_hash is None or cached_funding_hash is None:
        print("未找到缓存文件，需要完整计算统计...")
        need_recompute = True
    elif cached_paper_hash != data_hash or cached_funding_hash != data_hash:
        print("数据文件已更新，重新计算统计...")
        need_recompute = True
    else:
        need_recompute = False

    if not need_recompute:
        print("数据文件未变化，直接读取已有统计结果。")
        paper_stats = load_csv_with_hash(paper_csv)
        funding_stats = load_csv_with_hash(funding_csv)
        # 如果缓存的CSV缺少新增列，强制重新计算
        if not EXPECTED_PAPER_COLS.issubset(paper_stats.columns):
            print("检测到统计列有更新，重新计算文章统计...")
            people_df, talent_df, papers_df, vertical_df, horizontal_df = load_all_data(file_paths)
            teachers_df = filter_teachers(people_df)
            papers_cleaned = deduplicate_papers(papers_df)
            paper_stats = match_papers_for_teachers(teachers_df, papers_cleaned)
            tdm = _get_teacher_dept_map(teachers_df)
            paper_stats['院系'] = paper_stats['姓名'].map(tdm).fillna('')
            save_csv_with_hash(paper_stats, paper_csv, data_hash)
            print(f"文章统计已保存: {paper_csv}")
        else:
            # 仍需加载原始数据用于个人信息查询和散点图
            people_df, talent_df, _, _, _ = load_all_data(file_paths)
            # 补充院系列（如果缺失）
            if '院系' not in paper_stats.columns or '院系' not in funding_stats.columns:
                teachers_df = filter_teachers(people_df)
                tdm = _get_teacher_dept_map(teachers_df)
                if '院系' not in paper_stats.columns:
                    paper_stats['院系'] = paper_stats['姓名'].map(tdm).fillna('')
                    save_csv_with_hash(paper_stats, paper_csv, data_hash)
                if '院系' not in funding_stats.columns:
                    funding_stats['院系'] = funding_stats['姓名'].map(tdm).fillna('')
                    save_csv_with_hash(funding_stats, funding_csv, data_hash)
    else:
        # 1. 加载数据
        people_df, talent_df, papers_df, vertical_df, horizontal_df = load_all_data(file_paths)

        # 2. 筛选教学科研/工程实验教师
        teachers_df = filter_teachers(people_df)

        # 验证输入的教师存在
        all_teacher_names = set(teachers_df['姓名'].dropna().astype(str))
        for n in teacher_names:
            if n not in all_teacher_names:
                print(f"警告：教师 '{n}' 不在教学科研/工程实验教师名单中，但仍尝试统计。")

        # 3. 文章去重
        papers_cleaned = deduplicate_papers(papers_df)

        # 4. 文章匹配统计
        paper_stats = match_papers_for_teachers(teachers_df, papers_cleaned)
        tdm = _get_teacher_dept_map(teachers_df)
        paper_stats['院系'] = paper_stats['姓名'].map(tdm).fillna('')
        save_csv_with_hash(paper_stats, paper_csv, data_hash)
        print(f"文章统计已保存: {paper_csv}")

        # 5. 经费统计
        funding_stats = compute_funding_stats(teachers_df, vertical_df, horizontal_df)
        funding_stats['院系'] = funding_stats['姓名'].map(tdm).fillna('')
        save_csv_with_hash(funding_stats, funding_csv, data_hash)
        print(f"经费统计已保存: {funding_csv}")

    # 6. 合并统计数据（用于散点图）
    all_stats = paper_stats.merge(funding_stats, on='姓名', how='outer').fillna(0)

    # 7. 获取教师个人信息
    teacher_infos = {}
    for n in teacher_names:
        teacher_infos[n] = get_teacher_info(n, people_df, talent_df)

    print("=" * 60)
    print("生成报告和散点图...")

    # 8. 为每位教师生成个人报告和散点图
    for n in teacher_names:
        teacher_dir = output_path / n
        draw_scatter_single(all_stats, n, str(teacher_dir))
        generate_individual_report(n, paper_stats, funding_stats, teacher_infos[n], str(teacher_dir))
        print(f"  教师 {n} 的报告和散点图已生成。")

    # 9. 如果有多个教师，生成对比报告和对比散点图
    if len(teacher_names) > 1:
        compare_dir = output_path / "对比"
        draw_scatter_comparison(all_stats, teacher_names, str(compare_dir))
        generate_comparison_report(teacher_names, paper_stats, funding_stats, teacher_infos, str(compare_dir))
        print(f"  对比报告和散点图已生成。")

    print("=" * 60)
    print("全部完成！输出目录:", output_path)
    return str(output_path)


# ============================================================
# 院系统计功能
# ============================================================
VALID_DEPARTMENTS = ['天文', '物理', '核科学与技术']
DEPT_COL_NAME = '部门\n（行政人员部门暂为学院）'

DEPT_SCATTER_CONFIGS = [
    ('生涯总经费(万元)', '总文章数量'),
    ('生涯总经费(万元)', '第一署名单位文章数量'),
    ('生涯总经费(万元)', '通讯作者文章数量'),
    ('生涯总经费(万元)', '近五年文章数量'),
    ('生涯总经费(万元)', '近五年第一署名单位文章数量'),
    ('生涯总经费(万元)', '近五年通讯作者文章数量'),
    ('近五年总经费(万元)', '总文章数量'),
    ('近五年总经费(万元)', '第一署名单位文章数量'),
    ('近五年总经费(万元)', '通讯作者文章数量'),
    ('近五年总经费(万元)', '近五年文章数量'),
    ('近五年总经费(万元)', '近五年第一署名单位文章数量'),
    ('近五年总经费(万元)', '近五年通讯作者文章数量'),
]

DEPT_COLOR_MAP = {
    '天文': '#1f77b4',
    '物理': '#ff7f0e',
    '核科学与技术': '#2ca02c',
}

DEPT_COLORS = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
]


def _get_dept_talent_info(teachers_df, talent_df):
    """获取每位教师的院系（部门）和人才（特聘）信息，只保留天文/物理/核科学与技术"""
    talent_names = set()
    if talent_df is not None:
        for _, row in talent_df.iterrows():
            n = row.get('姓名', '')
            if pd.notna(n) and str(n).strip():
                talent_names.add(str(n).strip())

    info = {}
    for _, row in teachers_df.iterrows():
        name = str(row['姓名'])
        dept = row.get(DEPT_COL_NAME, '')
        if pd.isna(dept) or str(dept).strip() in ('', 'nan'):
            continue
        dept = str(dept).strip()
        if dept not in VALID_DEPARTMENTS:
            continue
        info[name] = {
            '院系': dept,
            '是否特聘': name in talent_names,
        }
    return info


def _get_teacher_dept_map(teachers_df):
    """构建教师姓名到院系的映射，只保留三个系"""
    teacher_dept = {}
    for _, row in teachers_df.iterrows():
        name = str(row['姓名'])
        dept = row.get(DEPT_COL_NAME, '')
        if pd.isna(dept) or str(dept).strip() in ('', 'nan'):
            continue
        dept = str(dept).strip()
        if dept in VALID_DEPARTMENTS:
            teacher_dept[name] = dept
    return teacher_dept


def draw_department_scatters(all_stats, dept_info, output_dir):
    """绘制散点图：每个系单独高亮图 + 三系总览图，共 12 × (3+1) = 48 张"""
    os.makedirs(output_dir, exist_ok=True)
    stats = all_stats.copy()
    # 只保留三个系的教师
    stats = stats[stats['姓名'].isin(dept_info)]
    stats['院系'] = stats['姓名'].map(lambda n: dept_info[n]['院系'])

    departments = VALID_DEPARTMENTS
    total_charts = 0

    for x_col, y_col in DEPT_SCATTER_CONFIGS:
        # --- 每个系的单独图（本系彩色+标注姓名，其他系灰色背景） ---
        for dept in departments:
            fig, ax = plt.subplots(figsize=(14, 10))
            other = stats[stats['院系'] != dept]
            ax.scatter(other[x_col], other[y_col],
                       c='lightgray', s=25, alpha=0.5, edgecolors='gray',
                       linewidths=0.3, label='其他系', zorder=1)

            sub = stats[stats['院系'] == dept]
            color = DEPT_COLOR_MAP[dept]
            ax.scatter(sub[x_col], sub[y_col], c=color, s=70, alpha=0.85,
                       edgecolors='black', linewidths=0.5, label=dept, zorder=2)
            for _, r in sub.iterrows():
                ax.annotate(r['姓名'], (r[x_col], r[y_col]),
                            fontsize=7, alpha=0.8, xytext=(4, 4),
                            textcoords='offset points',
                            color=color, fontweight='bold')

            ax.set_xlabel(x_col, fontsize=12)
            ax.set_ylabel(y_col, fontsize=12)
            ax.set_title(f'{dept}系 — {y_col} vs {x_col}', fontsize=14)
            ax.legend(loc='upper left', fontsize=10)
            ax.grid(True, alpha=0.3)
            plt.tight_layout()
            fname = f"{dept}_{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
            fig.savefig(os.path.join(output_dir, fname), dpi=150, bbox_inches='tight')
            plt.close(fig)
            total_charts += 1

        # --- 三系总览图（三种颜色区分） ---
        fig, ax = plt.subplots(figsize=(14, 10))
        for dept in departments:
            sub = stats[stats['院系'] == dept]
            color = DEPT_COLOR_MAP[dept]
            ax.scatter(sub[x_col], sub[y_col], c=color, s=50, alpha=0.7,
                       edgecolors='gray', linewidths=0.3, label=dept, zorder=2)
            for _, r in sub.iterrows():
                ax.annotate(r['姓名'], (r[x_col], r[y_col]),
                            fontsize=5, alpha=0.6, xytext=(3, 3),
                            textcoords='offset points')
        ax.set_xlabel(x_col, fontsize=12)
        ax.set_ylabel(y_col, fontsize=12)
        ax.set_title(f'{y_col} vs {x_col}（三系总览）', fontsize=14)
        ax.legend(loc='upper left', fontsize=10)
        ax.grid(True, alpha=0.3)
        plt.tight_layout()
        fname = f"总览_{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
        fig.savefig(os.path.join(output_dir, fname), dpi=150, bbox_inches='tight')
        plt.close(fig)
        total_charts += 1

    print(f"  散点图已保存: {total_charts} 张")


def draw_talent_comparison(all_stats, dept_info, output_dir):
    """绘制特聘 vs 非特聘对比图（文章和经费），只考虑三个系"""
    os.makedirs(output_dir, exist_ok=True)
    stats = all_stats.copy()
    stats = stats[stats['姓名'].isin(dept_info)]
    stats['类别'] = stats['姓名'].map(
        lambda n: '特聘' if dept_info[n]['是否特聘'] else '非特聘')

    talent = stats[stats['类别'] == '特聘']
    non_talent = stats[stats['类别'] == '非特聘']

    # --- 文章对比柱状图 ---
    paper_metrics = ['总文章数量', '近五年文章数量', '通讯作者文章数量', '近五年通讯作者文章数量']
    fig, ax = plt.subplots(figsize=(12, 6))
    x = list(range(len(paper_metrics)))
    w = 0.25
    means_all = [stats[m].mean() for m in paper_metrics]
    means_t = [talent[m].mean() for m in paper_metrics] if len(talent) > 0 else [0] * len(paper_metrics)
    means_nt = [non_talent[m].mean() for m in paper_metrics] if len(non_talent) > 0 else [0] * len(paper_metrics)

    for bars in [
        ax.bar([i - w for i in x], means_all, w, label='全院', color='gray', alpha=0.8),
        ax.bar(x, means_t, w, label='特聘教师', color='red', alpha=0.8),
        ax.bar([i + w for i in x], means_nt, w, label='非特聘教师', color='steelblue', alpha=0.8),
    ]:
        for bar in bars:
            h = bar.get_height()
            if h > 0:
                ax.text(bar.get_x() + bar.get_width() / 2, h, f'{h:.1f}',
                        ha='center', va='bottom', fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(paper_metrics, rotation=15, ha='right')
    ax.set_ylabel('人均文章数量')
    ax.set_title('特聘教师 vs 非特聘教师 — 人均文章数量对比')
    ax.legend()
    ax.grid(True, alpha=0.3, axis='y')
    plt.tight_layout()
    fig.savefig(os.path.join(output_dir, '特聘vs非特聘_文章对比.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)

    # --- 经费对比柱状图 ---
    funding_metrics = ['生涯总经费(万元)', '近五年总经费(万元)']
    fig, ax = plt.subplots(figsize=(8, 6))
    x = list(range(len(funding_metrics)))
    means_all = [stats[m].mean() for m in funding_metrics]
    means_t = [talent[m].mean() for m in funding_metrics] if len(talent) > 0 else [0] * len(funding_metrics)
    means_nt = [non_talent[m].mean() for m in funding_metrics] if len(non_talent) > 0 else [0] * len(funding_metrics)

    for bars in [
        ax.bar([i - w for i in x], means_all, w, label='全院', color='gray', alpha=0.8),
        ax.bar(x, means_t, w, label='特聘教师', color='red', alpha=0.8),
        ax.bar([i + w for i in x], means_nt, w, label='非特聘教师', color='steelblue', alpha=0.8),
    ]:
        for bar in bars:
            h = bar.get_height()
            if h > 0:
                ax.text(bar.get_x() + bar.get_width() / 2, h, f'{h:.1f}',
                        ha='center', va='bottom', fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(funding_metrics)
    ax.set_ylabel('人均经费（万元）')
    ax.set_title('特聘教师 vs 非特聘教师 — 人均经费对比')
    ax.legend()
    ax.grid(True, alpha=0.3, axis='y')
    plt.tight_layout()
    fig.savefig(os.path.join(output_dir, '特聘vs非特聘_经费对比.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)

    print("  特聘对比图已保存")
    return {
        'talent_count': len(talent),
        'non_talent_count': len(non_talent),
        'talent_paper_mean': talent['近五年文章数量'].mean() if len(talent) > 0 else 0,
        'non_talent_paper_mean': non_talent['近五年文章数量'].mean() if len(non_talent) > 0 else 0,
        'talent_funding_mean': talent['近五年总经费(万元)'].mean() if len(talent) > 0 else 0,
        'non_talent_funding_mean': non_talent['近五年总经费(万元)'].mean() if len(non_talent) > 0 else 0,
    }


def compute_yearly_funding_by_dept(teachers_df, vertical_df, horizontal_df):
    """按院系和年份统计经费（近五年），只考虑三个系"""
    five_year_start = current_year - RECENT_YEARS + 1
    years = list(range(five_year_start, current_year + 1))

    teacher_dept = _get_teacher_dept_map(teachers_df)
    teacher_names_set = set(teacher_dept.keys())

    v = vertical_df.copy()
    v['_year'] = v['立项日期'].apply(extract_date_year)
    v['_funding'] = v['批准经费'].apply(clean_funding)
    v = v[v['负责人'].isin(teacher_names_set)].copy()
    v['院系'] = v['负责人'].map(teacher_dept)

    h = horizontal_df.copy()
    h['_year'] = h['立项日期'].apply(extract_date_year)
    h['_funding'] = h['批准经费'].apply(clean_funding)
    h = h[h['负责人'].isin(teacher_names_set)].copy()
    h['院系'] = h['负责人'].map(teacher_dept)

    departments = VALID_DEPARTMENTS
    results = []
    for year in years:
        for dept in departments + ['总计']:
            if dept == '总计':
                vy = v[v['_year'] == year]
                hy = h[h['_year'] == year]
            else:
                vy = v[(v['_year'] == year) & (v['院系'] == dept)]
                hy = h[(h['_year'] == year) & (h['院系'] == dept)]
            vf = vy['_funding'].sum()
            hf = hy['_funding'].sum()
            results.append({
                '年份': year, '院系': dept,
                '纵向经费': round(vf, 2), '横向经费': round(hf, 2),
                '总经费': round(vf + hf, 2),
            })
    return pd.DataFrame(results)


def compute_yearly_papers_by_dept(teachers_df, papers_df):
    """按院系和年份统计文章数量（使用成果归属学者快速匹配），只考虑三个系"""
    five_year_start = current_year - RECENT_YEARS + 1
    years = list(range(five_year_start, current_year + 1))

    teacher_dept = _get_teacher_dept_map(teachers_df)

    papers = papers_df.copy()
    papers['_year'] = papers['年'].apply(extract_year)
    scholar_vals = papers['成果归属学者'].fillna('').astype(str).tolist()
    year_vals = papers['_year'].tolist()

    dept_year_counts = {}
    total_year_counts = {}

    for idx in range(len(papers)):
        yr = year_vals[idx]
        if yr is None or yr < five_year_start or yr > current_year:
            continue
        scholar_str = scholar_vals[idx]
        matched_depts = set()
        for tname, dept in teacher_dept.items():
            if tname in scholar_str:
                matched_depts.add(dept)
        for dept in matched_depts:
            dept_year_counts[(dept, yr)] = dept_year_counts.get((dept, yr), 0) + 1
        if matched_depts:
            total_year_counts[yr] = total_year_counts.get(yr, 0) + 1

    departments = VALID_DEPARTMENTS
    results = []
    for year in years:
        for dept in departments:
            results.append({
                '年份': year, '院系': dept,
                '文章数量': dept_year_counts.get((dept, year), 0),
            })
        results.append({
            '年份': year, '院系': '总计',
            '文章数量': total_year_counts.get(year, 0),
        })
    return pd.DataFrame(results)


def draw_yearly_funding_chart(yearly_funding, output_dir):
    """绘制近五年经费变化折线图（总经费/纵向/横向，分院系+总计）"""
    os.makedirs(output_dir, exist_ok=True)
    groups = list(yearly_funding['院系'].unique())
    # 确保总计在最后
    if '总计' in groups:
        groups.remove('总计')
    groups = sorted(groups) + ['总计']

    color_map = {d: DEPT_COLOR_MAP.get(d, DEPT_COLORS[i % len(DEPT_COLORS)])
                 for i, d in enumerate(groups) if d != '总计'}
    color_map['总计'] = 'black'

    fig, axes = plt.subplots(1, 3, figsize=(22, 7), sharey=False)
    for ax, metric, title in zip(axes,
            ['总经费', '纵向经费', '横向经费'],
            ['总经费变化', '纵向经费变化', '横向经费变化']):
        for dept in groups:
            sub = yearly_funding[yearly_funding['院系'] == dept].sort_values('年份')
            if dept == '总计':
                style = dict(linewidth=3, marker='s', markersize=8)
            else:
                style = dict(linewidth=1.5, marker='o', markersize=5)
            ax.plot(sub['年份'].astype(int), sub[metric], label=dept,
                    color=color_map.get(dept, 'gray'), **style)
        ax.set_title(title, fontsize=13)
        ax.set_xlabel('年份')
        ax.set_ylabel('经费（万元）')
        ax.grid(True, alpha=0.3)
        ax.legend(fontsize=7, ncol=2)
        years = sorted(yearly_funding['年份'].unique())
        ax.set_xticks(years)
        ax.set_xticklabels([str(int(y)) for y in years])

    fig.suptitle('近五年经费变化趋势（分院系 + 总计）', fontsize=15, fontweight='bold')
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(os.path.join(output_dir, '近五年经费变化趋势.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print("  经费趋势图已保存")


def draw_yearly_paper_chart(yearly_papers, output_dir):
    """绘制近五年文章数量变化折线图（分院系+总计）"""
    os.makedirs(output_dir, exist_ok=True)
    groups = list(yearly_papers['院系'].unique())
    if '总计' in groups:
        groups.remove('总计')
    groups = sorted(groups) + ['总计']

    color_map = {d: DEPT_COLOR_MAP.get(d, DEPT_COLORS[i % len(DEPT_COLORS)])
                 for i, d in enumerate(groups) if d != '总计'}
    color_map['总计'] = 'black'

    fig, ax = plt.subplots(figsize=(12, 7))
    for dept in groups:
        sub = yearly_papers[yearly_papers['院系'] == dept].sort_values('年份')
        if dept == '总计':
            style = dict(linewidth=3, marker='s', markersize=8)
        else:
            style = dict(linewidth=1.5, marker='o', markersize=5)
        ax.plot(sub['年份'].astype(int), sub['文章数量'], label=dept,
                color=color_map.get(dept, 'gray'), **style)

    ax.set_title('近五年文章数量变化趋势（分院系 + 总计）', fontsize=15, fontweight='bold')
    ax.set_xlabel('年份')
    ax.set_ylabel('文章数量')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=8, ncol=2)
    years = sorted(yearly_papers['年份'].unique())
    ax.set_xticks(years)
    ax.set_xticklabels([str(int(y)) for y in years])
    plt.tight_layout()
    fig.savefig(os.path.join(output_dir, '近五年文章数量变化趋势.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print("  文章趋势图已保存")


def generate_department_report(all_stats, dept_info, comparison,
                               yearly_funding, yearly_papers, output_dir):
    """生成院系统计 Markdown 报告"""
    os.makedirs(output_dir, exist_ok=True)

    lines = []
    lines.append(f"# 院系整体统计报告\n")
    lines.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    lines.append(f"近五年范围：{current_year - RECENT_YEARS + 1}—{current_year}\n")
    lines.append(f"统计院系：{'、'.join(VALID_DEPARTMENTS)}\n")

    # 基本概况
    total_teachers = len(all_stats)
    lines.append("## 基本概况\n")
    lines.append("| 指标 | 数值 |")
    lines.append("|------|------|")
    lines.append(f"| 教师总数（三系） | {total_teachers} |")
    lines.append(f"| 特聘教师数 | {comparison['talent_count']} |")
    lines.append(f"| 非特聘教师数 | {comparison['non_talent_count']} |")
    lines.append(f"| 全院人均总文章数 | {all_stats['总文章数量'].mean():.2f} |")
    lines.append(f"| 全院人均近五年文章数 | {all_stats['近五年文章数量'].mean():.2f} |")
    lines.append(f"| 全院人均生涯总经费（万元） | {all_stats['生涯总经费(万元)'].mean():.2f} |")
    lines.append(f"| 全院人均近五年经费（万元） | {all_stats['近五年总经费(万元)'].mean():.2f} |")
    lines.append("")

    # 分系概况
    stats_with_dept = all_stats.copy()
    stats_with_dept['院系'] = stats_with_dept['姓名'].map(lambda n: dept_info.get(n, {}).get('院系', ''))
    for dept in VALID_DEPARTMENTS:
        sub = stats_with_dept[stats_with_dept['院系'] == dept]
        if len(sub) > 0:
            lines.append(f"### {dept}系\n")
            lines.append(f"- 教师人数：{len(sub)}")
            lines.append(f"- 人均总文章数：{sub['总文章数量'].mean():.2f}")
            lines.append(f"- 人均近五年文章数：{sub['近五年文章数量'].mean():.2f}")
            lines.append(f"- 人均生涯总经费：{sub['生涯总经费(万元)'].mean():.2f} 万元")
            lines.append(f"- 人均近五年经费：{sub['近五年总经费(万元)'].mean():.2f} 万元")
            lines.append("")

    # 特聘对比
    lines.append("## 特聘 vs 非特聘 对比\n")
    lines.append("| 指标 | 特聘教师 | 非特聘教师 |")
    lines.append("|------|---------|----------|")
    lines.append(f"| 人数 | {comparison['talent_count']} | {comparison['non_talent_count']} |")
    lines.append(f"| 人均近五年文章数 | {comparison['talent_paper_mean']:.2f} | {comparison['non_talent_paper_mean']:.2f} |")
    lines.append(f"| 人均近五年经费（万元） | {comparison['talent_funding_mean']:.2f} | {comparison['non_talent_funding_mean']:.2f} |")
    lines.append("")
    lines.append("### 文章对比\n")
    lines.append("![特聘vs非特聘_文章对比](特聘vs非特聘_文章对比.png)\n")
    lines.append("### 经费对比\n")
    lines.append("![特聘vs非特聘_经费对比](特聘vs非特聘_经费对比.png)\n")

    # 散点图
    lines.append("## 散点图分布\n")
    for x_col, y_col in DEPT_SCATTER_CONFIGS:
        # 总览图
        fname_all = f"总览_{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
        lines.append(f"### {y_col} vs {x_col}（三系总览）\n")
        lines.append(f"![三系总览]({fname_all})\n")
        # 各系图
        for dept in VALID_DEPARTMENTS:
            fname = f"{dept}_{y_col}_vs_{x_col}.png".replace('/', '_').replace('(', '').replace(')', '')
            lines.append(f"### {y_col} vs {x_col}（{dept}系）\n")
            lines.append(f"![{dept}系]({fname})\n")

    # 经费趋势
    lines.append("## 近五年经费变化趋势\n")
    lines.append("![近五年经费变化趋势](近五年经费变化趋势.png)\n")
    lines.append("### 经费数据明细（总经费，万元）\n")
    pivot_f = yearly_funding.pivot_table(
        index='院系', columns='年份', values='总经费', aggfunc='sum').fillna(0)
    year_cols = sorted(pivot_f.columns)
    header = "| 院系 | " + " | ".join(str(int(y)) for y in year_cols) + " |"
    sep = "|------|" + "|".join(["------"] * len(year_cols)) + "|"
    lines.append(header)
    lines.append(sep)
    for dept in pivot_f.index:
        vals = " | ".join(f"{pivot_f.loc[dept, y]:.2f}" for y in year_cols)
        lines.append(f"| {dept} | {vals} |")
    lines.append("")

    # 文章趋势
    lines.append("## 近五年文章数量变化趋势\n")
    lines.append("![近五年文章数量变化趋势](近五年文章数量变化趋势.png)\n")
    lines.append("### 文章数据明细\n")
    pivot_p = yearly_papers.pivot_table(
        index='院系', columns='年份', values='文章数量', aggfunc='sum').fillna(0)
    year_cols_p = sorted(pivot_p.columns)
    header = "| 院系 | " + " | ".join(str(int(y)) for y in year_cols_p) + " |"
    sep = "|------|" + "|".join(["------"] * len(year_cols_p)) + "|"
    lines.append(header)
    lines.append(sep)
    for dept in pivot_p.index:
        vals = " | ".join(str(int(pivot_p.loc[dept, y])) for y in year_cols_p)
        lines.append(f"| {dept} | {vals} |")
    lines.append("")

    report_path = os.path.join(output_dir, "院系统计报告.md")
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"  报告已生成: {report_path}")
    md_to_docx(report_path)


def run_department_stats(file_paths=None, output_path=None):
    """运行院系整体统计分析"""
    if output_path is not None:
        output_path = Path(output_path)
    else:
        output_path = OUTPUT_PATH

    dept_output = output_path / "院系统计"
    os.makedirs(dept_output, exist_ok=True)

    print("=" * 60)
    print("开始院系整体统计分析")
    print(f"统计院系：{'、'.join(VALID_DEPARTMENTS)}")
    print(f"当前年份：{current_year}，近五年范围：{current_year - RECENT_YEARS + 1}-{current_year}")
    print("=" * 60)

    # 1. 加载数据
    data_hash = compute_data_hash(file_paths)
    cache_path = output_path / "cache"
    os.makedirs(cache_path, exist_ok=True)
    paper_csv = cache_path / "文章统计.csv"
    funding_csv = cache_path / "经费统计.csv"

    people_df, talent_df, papers_df, vertical_df, horizontal_df = load_all_data(file_paths)
    teachers_df = filter_teachers(people_df)
    papers_cleaned = deduplicate_papers(papers_df)

    # 构建教师院系映射
    teacher_dept_map = _get_teacher_dept_map(teachers_df)

    # 2. 文章/经费统计（利用缓存）
    cached_paper_hash = read_hash_from_csv(paper_csv)
    cached_funding_hash = read_hash_from_csv(funding_csv)
    EXPECTED_PAPER_COLS = {'第一署名单位文章数量', '近五年第一署名单位文章数量',
                           '通讯作者文章数量', '近五年通讯作者文章数量'}

    if cached_paper_hash == data_hash and cached_funding_hash == data_hash:
        print("使用缓存的统计数据...")
        paper_stats = load_csv_with_hash(paper_csv)
        funding_stats = load_csv_with_hash(funding_csv)
        if not EXPECTED_PAPER_COLS.issubset(paper_stats.columns):
            paper_stats = match_papers_for_teachers(teachers_df, papers_cleaned)
            paper_stats['院系'] = paper_stats['姓名'].map(teacher_dept_map).fillna('')
            save_csv_with_hash(paper_stats, paper_csv, data_hash)
        elif '院系' not in paper_stats.columns:
            paper_stats['院系'] = paper_stats['姓名'].map(teacher_dept_map).fillna('')
            save_csv_with_hash(paper_stats, paper_csv, data_hash)
        if '院系' not in funding_stats.columns:
            funding_stats['院系'] = funding_stats['姓名'].map(teacher_dept_map).fillna('')
            save_csv_with_hash(funding_stats, funding_csv, data_hash)
    else:
        print("计算统计数据（首次或数据已更新）...")
        paper_stats = match_papers_for_teachers(teachers_df, papers_cleaned)
        paper_stats['院系'] = paper_stats['姓名'].map(teacher_dept_map).fillna('')
        save_csv_with_hash(paper_stats, paper_csv, data_hash)
        funding_stats = compute_funding_stats(teachers_df, vertical_df, horizontal_df)
        funding_stats['院系'] = funding_stats['姓名'].map(teacher_dept_map).fillna('')
        save_csv_with_hash(funding_stats, funding_csv, data_hash)

    all_stats = paper_stats.merge(funding_stats, on='姓名', how='outer', suffixes=('', '_dup')).fillna(0)
    # 解决合并后可能出现的重复院系列
    if '院系_dup' in all_stats.columns:
        all_stats.drop(columns=['院系_dup'], inplace=True)

    dept_info = _get_dept_talent_info(teachers_df, talent_df)

    # 只保留三个系的教师统计
    dept_stats = all_stats[all_stats['姓名'].isin(dept_info)].copy()
    print(f"三系教师总数: {len(dept_stats)}（天文: {sum(1 for v in dept_info.values() if v['院系']=='天文')}，"
          f"物理: {sum(1 for v in dept_info.values() if v['院系']=='物理')}，"
          f"核科学与技术: {sum(1 for v in dept_info.values() if v['院系']=='核科学与技术')}）")

    # 3. 绘制图表
    print("生成散点图...")
    draw_department_scatters(dept_stats, dept_info, str(dept_output))

    print("生成特聘对比图...")
    comparison = draw_talent_comparison(dept_stats, dept_info, str(dept_output))

    print("计算并绘制年度经费趋势...")
    yearly_funding = compute_yearly_funding_by_dept(teachers_df, vertical_df, horizontal_df)
    draw_yearly_funding_chart(yearly_funding, str(dept_output))

    print("计算并绘制年度文章趋势...")
    yearly_papers = compute_yearly_papers_by_dept(teachers_df, papers_cleaned)
    draw_yearly_paper_chart(yearly_papers, str(dept_output))

    # 4. 保存数据和生成报告
    yearly_funding.to_csv(os.path.join(str(dept_output), '年度经费统计.csv'),
                          index=False, encoding='utf-8-sig')
    yearly_papers.to_csv(os.path.join(str(dept_output), '年度文章统计.csv'),
                         index=False, encoding='utf-8-sig')
    dept_stats.to_csv(os.path.join(str(dept_output), '教师综合统计.csv'),
                      index=False, encoding='utf-8-sig')

    print("生成统计报告...")
    generate_department_report(dept_stats, dept_info, comparison,
                               yearly_funding, yearly_papers, str(dept_output))

    print("=" * 60)
    print(f"院系统计完成！输出目录: {dept_output}")
    return str(dept_output)


if __name__ == "__main__":
    main()
