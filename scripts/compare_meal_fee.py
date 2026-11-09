#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
餐费两版比对（meal-fee-calc）

比对两份餐费结果，回答三件事：差多少、差在哪一笔、谁对。

用法：
    python3 compare_meal_fee.py \
      --a <A方表：xlsx/csv> --b <B方表：xlsx/csv> \
      --out <输出目录> \
      [--ot <加班记录.xlsx>] [--region 广东] \
      [--a-name 本方案] [--b-name 对方] \
      [--adjust-workdays "2026-09-20,2026-10-10"] \
      [--shift-end 17:15]

支持的输入形态（自动探测，无需指定）：
    1. 逐笔明细：有 工号 + 日期 + 餐别 + 金额 列
    2. 按人汇总：有 工号 + 金额 列
    3. 汇总 + 差异页混合表（如区域回传的核对表）：汇总页取最终金额，差异页作逐笔线索

输出（<out>/餐费比对_<A名>_vs_<B名>.xlsx）：
    比对结论 / 逐人差异 / 逐笔线索

原则：
    - 金额列只在「最终应付/应扣」白名单里选，"核对后修改""差额""刷卡金额"一律不当金额
    - 逐人比对是权威口径；逐笔只在双方（或至少一方）有逐笔数据时输出，用于定位
    - 归因规则读 references/差异归因规则.json；命中不了标"需人工判定"，不硬套
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime, time, timedelta

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(HERE)
RULE_JSON = os.path.join(SKILL_ROOT, 'references', '差异归因规则.json')

SLOTS = {'早餐': (5 * 60 + 30, 9 * 60), '午餐': (10 * 60 + 30, 14 * 60 + 30),
         '晚餐': (15 * 60 + 30, 20 * 60 + 30)}

# 金额列白名单：只有"最终应扣/应付"才算数
AMOUNT_COLS = ['员工应付', '应扣餐费', '应扣', '员工自付', '自付金额', '员工实际应付',
               '实际应付', '应付金额', '金额合计']
# 绝对不能当金额的列
AMOUNT_EXCLUDE = ['核对后修改', '差额', '刷卡金额', '打卡总金额', '餐标金额',
                  '开餐补助金额', '实际金额', '抵扣金额', '加班时长']

ID_COLS = ['工号', '员工工号', '工号ID', '人员ID']
NAME_COLS = ['姓名', '员工姓名']
DATE_COLS = ['打卡日期', '日期', '用餐日期']
MEAL_COLS = ['开餐餐次', '餐次', '餐别']
TIME_COLS = ['打卡时间', '用餐时间', '刷卡时间']


# ---------------- 工具 ----------------
def norm_id(x):
    """工号规范化：去空格、去 .0、纯数字去前导零、字母转大写"""
    if pd.isna(x):
        return None
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return str(int(s)) if s.isdigit() else s.upper()


def find_col(df, cands, fuzzy=True):
    cols = {str(c).strip(): c for c in df.columns}
    for c in cands:
        if c in cols:
            return cols[c]
    if fuzzy:
        for c in cands:
            for k, v in cols.items():
                if c in k:
                    return v
    return None


def to_min(t):
    """'08:31:55' / '2026-09-29 08:31:55' → 分钟数"""
    if pd.isna(t):
        return None
    m = re.search(r'(\d{1,2}):(\d{2})', str(t))
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def is_result_table(df):
    """是否像"逐人结果"：有工号且有金额列"""
    return find_col(df, ID_COLS) is not None and find_col(df, AMOUNT_COLS) is not None


def is_detail_table(df):
    """是否像"逐笔明细"：有工号、日期、餐别"""
    return (find_col(df, ID_COLS) is not None
            and find_col(df, DATE_COLS) is not None
            and find_col(df, MEAL_COLS) is not None)


def pick_amount_col(df):
    """在白名单里挑最权威的金额列"""
    for c in AMOUNT_COLS:
        if c in [str(x).strip() for x in df.columns]:
            return c
    return None


def load_one(path):
    """读一个文件，返回 {'per': 逐人表, 'per_src':, 'detail': 逐笔表, 'detail_src':, 'sheets': []}"""
    out = {'per': None, 'per_src': None, 'detail': None, 'detail_src': None, 'sheets': []}
    if path.lower().endswith('.csv'):
        df = pd.read_csv(path, dtype={'工号': str, '员工工号': str})
        out['sheets'] = ['csv']
        if is_detail_table(df):
            out['detail'], out['detail_src'] = df, 'csv'
            out['per'] = aggregate_detail(df)
            out['per_src'] = 'csv（由逐笔汇总）'
        elif is_result_table(df):
            out['per'], out['per_src'] = df, 'csv'
        return out

    xl = pd.ExcelFile(path)
    out['sheets'] = xl.sheet_names
    detail_cands, per_cands, clue_cands = [], [], []

    for sh in xl.sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=sh, dtype={'工号': str, '员工工号': str})
        except Exception:
            continue
        if df.empty:
            continue
        # 逐笔：有日期+餐别 且 行数明显多
        if is_detail_table(df) and len(df) > 5:
            if '差异' in sh or '线索' in sh:
                clue_cands.append((sh, df))
            else:
                detail_cands.append((sh, df))
        elif is_result_table(df):
            per_cands.append((sh, df))

    # 逐人表：优先 sheet 名含"汇总"，其次行数最少的（人数一般少于明细）
    if per_cands:
        prefer = [x for x in per_cands if '汇总' in x[0]]
        chosen = prefer[0] if prefer else min(per_cands, key=lambda x: len(x[1]))
        out['per'], out['per_src'] = chosen[1], chosen[0]
    # 逐笔表
    if detail_cands:
        chosen = max(detail_cands, key=lambda x: len(x[1]))
        out['detail'], out['detail_src'] = chosen[1], chosen[0]
    # 逐笔线索（对方只给了部分笔，如"差异"页）
    out['clue'] = clue_cands[0][1] if clue_cands else None
    out['clue_src'] = clue_cands[0][0] if clue_cands else None
    return out


def aggregate_detail(df):
    """逐笔 → 逐人（按工号汇总自付金额）"""
    idc = find_col(df, ID_COLS)
    amtc = pick_amount_col(df) or find_col(df, ['员工自付', '自付金额'])
    nmc = find_col(df, NAME_COLS)
    if not (idc and amtc):
        return None
    t = df[[idc, nmc, amtc]].copy() if nmc else df[[idc, amtc]].copy()
    t.columns = ['工号', '姓名', '应扣餐费'] if nmc else ['工号', '应扣餐费']
    t['工号'] = t['工号'].map(norm_id)
    t['应扣餐费'] = pd.to_numeric(t['应扣餐费'], errors='coerce').fillna(0)
    return t.groupby('工号', as_index=False)['应扣餐费'].sum()


def norm_per(df):
    """逐人表统一成 [工号, 姓名, 金额]"""
    idc = find_col(df, ID_COLS)
    amtc = pick_amount_col(df)
    nmc = find_col(df, NAME_COLS)
    if not idc or not amtc:
        return None
    res = pd.DataFrame({
        '工号': df[idc].map(norm_id),
        '姓名': df[nmc].astype(str).str.strip() if nmc else '',
        '金额': pd.to_numeric(df[amtc], errors='coerce').fillna(0).round(2),
    })
    res = res.dropna(subset=['工号'])
    return res.groupby('工号', as_index=False).agg(姓名=('姓名', 'first'), 金额=('金额', 'sum'))


def load_ot(path):
    """加班记录 → {(工号, 归属日): {'mins':, 'segs': []}}"""
    if not path or not os.path.exists(path):
        return {}
    raw = pd.read_excel(path, header=None, nrows=2)
    r0 = ' '.join(str(x) for x in raw.iloc[0])
    r1 = ' '.join(str(x) for x in raw.iloc[1])
    header = 1 if ('加班日期' in r1 and '员工工号' in r1) else 0
    ot = pd.read_excel(path, header=header)
    cmap = {}
    for c in ot.columns:
        cs = str(c).strip()
        if '员工工号' in cs or cs == 'StaffId-StaffCode':
            cmap['id'] = c
        elif '加班日期' in cs or cs == 'OverTimeDate':
            cmap['date'] = c
        elif '实际打卡加班时间' in cs or cs == 'ActualStartDateTime':
            cmap['st'] = c
        elif '实际打卡结束时间' in cs or cs == 'ActualStopDateTime':
            cmap['en'] = c
        elif '实际加班时长' in cs or cs == 'ActualDurationIncludeUnit':
            cmap['mins'] = c
        elif '单据类型' in cs or cs == 'DocumentType':
            cmap['type'] = c
    ot = ot.rename(columns={v: k for k, v in cmap.items()})
    if 'type' in ot.columns:
        ot = ot[ot['type'].astype(str).str.strip() != '撤销']
    ot['id'] = ot['id'].astype(str).str.strip()
    ot['st'] = pd.to_datetime(ot.get('st'), errors='coerce')
    ot['en'] = pd.to_datetime(ot.get('en'), errors='coerce')
    m = {}

    def mins_of(v):
        n = re.search(r'(\d+(?:\.\d+)?)', str(v))
        return int(float(n.group(1))) if n else 0

    for _, r in ot.iterrows():
        attr = r['en'].date() if pd.notna(r['en']) else None
        if attr is None and 'date' in ot.columns and pd.notna(r['date']):
            attr = pd.to_datetime(r['date']).date()
        if attr is None:
            continue
        e = m.setdefault((str(r['id']).strip(), attr), {'mins': 0, 'segs': []})
        e['mins'] += mins_of(r.get('mins', 0))
        if pd.notna(r['st']) and pd.notna(r['en']):
            e['segs'].append((r['st'], r['en']))
    return m


def covered(segs, day):
    """加班段覆盖的餐别（真实时段重叠）"""
    res = set()
    if not segs:
        return res
    for name, (s, e) in SLOTS.items():
        ws = datetime.combine(day, time(0)) + timedelta(minutes=s)
        we = datetime.combine(day, time(0)) + timedelta(minutes=e)
        for st, en in segs:
            if st <= we and en >= ws:
                res.add(name)
                break
    return res


def load_rules():
    """归因规则：优先读 references/差异归因规则.json，缺省用内置兜底"""
    if os.path.exists(RULE_JSON):
        try:
            with open(RULE_JSON, encoding='utf-8') as f:
                return json.load(f).get('patterns', [])
        except Exception:
            pass
    return []


def match_rule(patterns, feat):
    """按特征匹配归因规则；feature 中每个键都必须满足（值可为列表=属于其中之一）"""
    for p in patterns:
        cond = p.get('feature') or {}
        hit = True
        for k, v in cond.items():
            fv = feat.get(k)
            if isinstance(v, list):
                if fv not in v:
                    hit = False
                    break
            elif fv != v:
                hit = False
                break
        if hit:
            return p
    return None


def responsibility(correct, a_pay, b_pay):
    """correct='应抵'/'应自付'/None；返回谁与现行规则不符"""
    if not correct:
        return '需人工判定'
    if b_pay is None:
        return '对方未提供该笔逐笔判定'
    a_deduct = a_pay is not None and a_pay <= 0.001
    b_deduct = b_pay <= 0.001
    if correct == '应抵':
        bad = []
        if not a_deduct:
            bad.append('A')
        if not b_deduct:
            bad.append('B')
        return '双方均符合' if not bad else f'{"、".join(bad)}方与现行规则不符'
    if correct == '应自付':
        bad = []
        if a_deduct:
            bad.append('A')
        if b_deduct:
            bad.append('B')
        return '双方均符合' if not bad else f'{"、".join(bad)}方与现行规则不符'
    return '需人工判定'


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser(description='餐费两版比对')
    ap.add_argument('--a', required=True, help='A方表（xlsx/csv）')
    ap.add_argument('--b', required=True, help='B方表（xlsx/csv）')
    ap.add_argument('--out', required=True, help='输出目录')
    ap.add_argument('--ot', default=None, help='加班记录（可选，用于判定加班是否覆盖餐别）')
    ap.add_argument('--shift-end', default='17:15', help='默认下班时刻，用于识别"下班后刷卡"')
    ap.add_argument('--a-name', default='A方')
    ap.add_argument('--b-name', default='B方')
    ap.add_argument('--adjust-workdays', default='2026-09-20,2026-10-10',
                    help='调休上班日（逗号分隔）')
    args = ap.parse_args()

    for p in (args.a, args.b):
        if not os.path.exists(p):
            sys.exit(f'文件不存在：{p}')

    os.makedirs(args.out, exist_ok=True)
    OT = load_ot(args.ot)
    end_min = to_min(args.shift_end) or (17 * 60 + 15)
    adj = {pd.Timestamp(x.strip()).date() for x in args.adjust_workdays.split(',') if x.strip()}

    # ---- 读两份 ----
    A = load_one(args.a)
    B = load_one(args.b)
    if A['per'] is None or B['per'] is None:
        sys.exit(f'无法识别结果表：A={A["per_src"]}，B={B["per_src"]}。'
                 f'请确认表中含「工号」列与「应扣/应付金额」列。')

    pa = norm_per(A['per'])
    pb = norm_per(B['per'])
    if pa is None or pb is None:
        sys.exit('结果表缺少工号列或金额列，无法比对。')

    # ---- 逐人比对 ----
    j = pa.rename(columns={'姓名': f'{args.a_name}_姓名', '金额': 'A金额'}).set_index('工号').join(
        pb.rename(columns={'姓名': f'{args.b_name}_姓名', '金额': 'B金额'}).set_index('工号'),
        how='outer')
    j['A金额'] = j['A金额'].fillna(0).round(2)
    j['B金额'] = j['B金额'].fillna(0).round(2)
    j['差额'] = (j['B金额'] - j['A金额']).round(2)
    j['姓名'] = j[f'{args.a_name}_姓名'].fillna(j[f'{args.b_name}_姓名']).fillna('')
    j = j.reset_index()

    total_a = float(j['A金额'].sum())
    total_b = float(j['B金额'].sum())
    n_a, n_b = len(pa), len(pb)
    only_a = set(pa['工号']) - set(pb['工号'])
    only_b = set(pb['工号']) - set(pa['工号'])
    changed = j[j['差额'].abs() > 0.011].copy().sort_values('差额', ascending=False)

    # ---- 逐笔线索：把 A 的逐笔按 (工号, 日期, 餐别) 取出，补特征 ----
    detail_rows = []
    adet = A['detail']
    if adet is not None:
        idc = find_col(adet, ID_COLS)
        dc = find_col(adet, DATE_COLS)
        mc = find_col(adet, MEAL_COLS)
        tc = find_col(adet, TIME_COLS)
        payc = find_col(adet, ['员工自付', '自付金额', '员工应付'])
        swc = find_col(adet, ['刷卡金额', '打卡总金额'])
        rsnc = find_col(adet, ['判定', '判定说明', '原因'])
        adet['_id'] = adet[idc].map(norm_id)
        adet['_date'] = pd.to_datetime(adet[dc], errors='coerce')
        adet['_meal'] = adet[mc].astype(str).str.strip()
        # 同日同餐排序：第 2 笔及以后（规则 6.3 只有第一笔参与抵扣）
        adet = adet.sort_values(['_id', '_date', '_meal'] + ([tc] if tc else [])).reset_index(drop=True)
        adet['_seq'] = adet.groupby(['_id', '_date', '_meal']).cumcount() + 1
        # 当天同餐笔数
        grp_cnt = adet.groupby(['_id', '_date', '_meal'])['_id'].transform('size')
        adet['_dup'] = grp_cnt > 1
        # 当天该人实际刷卡的餐别集合（判断"是否只吃了一餐"）
        adet['_day'] = adet['_date'].dt.date
        day_meals = adet.groupby(['_id', '_day'])['_meal'].apply(lambda s: set(s)).to_dict()
        # 当天 A 方抵扣了几笔（判断是否超额）
        if payc:
            adet['_a_pay'] = pd.to_numeric(adet[payc], errors='coerce').fillna(0)
            day_deduct_cnt = adet[adet['_a_pay'] <= 0.001].groupby(['_id', '_day']).size().to_dict()

        B_clue = {}
        if B.get('clue') is not None:
            cdf = B['clue']
            cid = find_col(cdf, ID_COLS)
            cdc = find_col(cdf, DATE_COLS)
            cmc = find_col(cdf, MEAL_COLS)
            # 对方逐笔金额：优先"员工实际应付"（对方判定），其次白名单
            cpay = find_col(cdf, ['员工实际应付'])
            if cpay is None:
                cpay = pick_amount_col(cdf)
            if cid and cdc and cmc and cpay:
                for _, r in cdf.iterrows():
                    B_clue[(norm_id(r[cid]), pd.to_datetime(r[cdc], errors='coerce').date(),
                            str(r[cmc]).strip())] = float(pd.to_numeric(r[cpay], errors='coerce') or 0)

        # 对方当天各餐别的抵扣情况（用于识别"额度在午晚之间分配不同"）
        b_day_deduct = {}
        if B_clue:
            for (bid, bday, bmeal), bv in B_clue.items():
                if bv <= 0.001:
                    b_day_deduct.setdefault((bid, bday), set()).add(bmeal)

        # 逐笔特征 + 归因（只对差异人）
        chg_ids = set(changed['工号'])
        chg_map = changed.set_index('工号')['差额'].to_dict()
        for _, r in adet.iterrows():
            if r['_id'] not in chg_ids:
                continue
            day = r['_date'].date() if pd.notna(r['_date']) else None
            meal = r['_meal']
            t_min = to_min(r[tc]) if tc else None
            a_pay = round(float(r[payc] or 0), 2) if payc else None
            b_pay = B_clue.get((r['_id'], day, meal))
            p_diff = float(chg_map.get(r['_id'], 0))
            oti = OT.get((r['_id'], day), {'mins': 0, 'segs': []})
            cov = covered(oti['segs'], day) if oti['segs'] else set()
            weekend = (day.weekday() >= 5 and day not in adj) if day else None
            sw = float(r[swc] or 0) if swc else None
            reason_txt = str(r[rsnc]) if rsnc else ''

            # 线索类型：双方均有判定且不同=精确差异；对方无该笔=按差额金额+方向筛候选
            if b_pay is not None and a_pay is not None and abs(b_pay - a_pay) > 0.011:
                clue_type, is_cand = '精确差异（双方均有判定）', False
            elif b_pay is not None:
                continue  # 双方判定一致，不占篇幅
            else:
                # 对方无该笔逐笔：此处不输出，循环后按「逐人净额校验」统一给结论
                continue

            a_cnt = day_deduct_cnt.get((r['_id'], day), 0)
            b_cnt = len(b_day_deduct.get((r['_id'], day), set()))
            feat = {
                'meal': meal,
                'is_weekend': weekend,
                'swipe_after_shift_end': bool(t_min is not None and t_min >= end_min),
                'ot_mins_ge_60': oti['mins'] >= 60,
                'ot_mins_180_300': 180 <= oti['mins'] <= 300,
                'ot_covers_meal': (meal in cov) if cov else False,
                'same_meal_dup': bool(r['_dup']),
                'is_second_or_later': int(r['_seq']) >= 2,
                'only_meal_of_day': len(day_meals.get((r['_id'], day), set())) == 1,
                'a_deducted_two_or_more': a_cnt >= 2,
                'a_deduct_cnt': a_cnt,
                'b_deduct_cnt': b_cnt,
                'same_deduct_count': a_cnt == b_cnt,
                'a_reason_work_meal': '工作餐' in str(r[rsnc] if rsnc else ''),
                'clue_is_candidate': is_cand,
                'other_meal_b_deducted': len(b_day_deduct.get((r['_id'], day), set()) - {meal}) > 0,
                'max_deduct_more_than_one_today': b_cnt > 1,
                'a_deduct': (a_pay is not None and a_pay <= 0.001),
                'b_deduct': (b_pay is not None and b_pay <= 0.001),
            }
            rule = match_rule(load_rules(), feat)

            detail_rows.append({
                '线索类型': clue_type,
                '工号': r['_id'],
                '姓名': j.loc[j['工号'] == r['_id'], '姓名'].iloc[0] if (j['工号'] == r['_id']).any() else '',
                '日期': str(day) if day else '',
                '星期': '一二三四五六日'[day.weekday()] if day else '',
                '餐别': meal,
                '刷卡时间': str(r[tc])[:8] if tc and pd.notna(r[tc]) else '',
                '刷卡金额': float(r[swc] or 0) if swc else None,
                f'{args.a_name}判定金额': a_pay,
                f'{args.b_name}判定金额': b_pay,
                '本笔差额': round((b_pay - a_pay), 2) if b_pay is not None and a_pay is not None else None,
                '当日加班(分)': oti['mins'],
                '加班覆盖': '/'.join(sorted(cov)) if cov else '无',
                '同日同餐笔数>1': '是' if r['_dup'] else '否',
                f'{args.a_name}判定说明': str(r[rsnc]) if rsnc else '',
                '候选归因': (rule or {}).get('name', '需人工判定（未命中已归档模式）'),
                '归因类型': (rule or {}).get('type', ''),
                '规则结论': (rule or {}).get('correct', '') or '',
                '责任判定': responsibility((rule or {}).get('correct'), a_pay, b_pay)
                            if (rule and rule.get('correct')) else '需人工判定',
                '规则依据': (rule or {}).get('basis', ''),
            })

    # 逐人净额校验：已定位的精确差异之和，是否解释了该人的全部差额
    explained, seen_cnt = {}, {}
    listed = {(row['工号'], row['日期'], row['餐别']) for row in detail_rows if row['日期']}
    for row in detail_rows:
        pid0 = row['工号']
        explained[pid0] = explained.get(pid0, 0) + (row['本笔差额'] or 0)
        seen_cnt[pid0] = seen_cnt.get(pid0, 0) + 1
    for _, r in changed.iterrows():
        pid, diff = r['工号'], round(float(r['差额']), 2)
        exp_sum = round(explained.get(pid, 0), 2)
        gap = round(diff - exp_sum, 2)
        if abs(gap) <= 0.011:
            continue  # 差额已被逐笔解释完整
        # 找疑似笔：刷卡金额≈|gap|、本方判定方向与差额相反、当日加班覆盖该餐、且尚未列入
        suspect, suspect_day = '', ''
        if swc and payc and (adet['_id'] == pid).any():
            for _, x in adet[adet['_id'] == pid].iterrows():
                sw_v = float(x[swc] or 0)
                ap_v = float(x[payc] or 0)
                if not (abs(sw_v - abs(gap)) <= 0.011
                        and ((gap > 0 and ap_v <= 0.001) or (gap < 0 and ap_v > 0.001))):
                    continue
                d2 = pd.to_datetime(x['_date']).date() if pd.notna(x['_date']) else None
                if (pid, str(d2), x['_meal']) in listed:
                    continue  # 该笔已在精确差异中列过
                o2 = OT.get((pid, d2), {'mins': 0, 'segs': []})
                cv2 = covered(o2['segs'], d2) if o2['segs'] else set()
                if not (o2['mins'] >= 60 and x['_meal'] in cv2):
                    continue  # 无加班覆盖的笔不可能产生抵扣争议
                suspect_day = str(d2)
                suspect = f'{d2} {x["_meal"]}：刷卡{sw_v:g}元，{args.a_name}判{ap_v:g}元（{x[rsnc] if rsnc else ""}）'
                break
        detail_rows.append({
            '线索类型': '差额未解释（对方未提供逐笔或做了手工调整）',
            '工号': pid, '姓名': r['姓名'], '日期': suspect_day, '星期': '',
            '餐别': '', '刷卡时间': '', '刷卡金额': None,
            f'{args.a_name}判定金额': None, f'{args.b_name}判定金额': None,
            '本笔差额': gap,
            '当日加班(分)': None, '加班覆盖': '', '同日同餐笔数>1': '',
            f'{args.a_name}判定说明': suspect,
            '候选归因': (f'对方逐笔净额 {exp_sum:+.1f} 元与其汇总差额 {diff:+.1f} 元不一致，缺口 {gap:+.1f} 元无法归到具体笔'
                         if seen_cnt.get(pid, 0) else '对方未提供该人逐笔，差额无法对应到具体笔'),
            '归因类型': '内部不自洽' if seen_cnt.get(pid, 0) else '缺乏对方逐笔',
            '规则结论': '',
            '责任判定': '需人工判定',
            '规则依据': (f'汇总差额 {diff:+.1f} 元 vs 逐笔已解释 {exp_sum:+.1f} 元，缺口 {gap:+.1f} 元；'
                         '请对方补充逐笔明细或说明调整依据'),
        })

    det = pd.DataFrame(detail_rows)

    # ---- 输出 xlsx ----
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    hdr_fill = PatternFill('solid', fgColor='3370FF')
    hdr_font = Font(name='微软雅黑', bold=True, color='FFFFFF', size=10)
    body_font = Font(name='微软雅黑', size=10)
    thin = Side(style='thin', color='B8C7D9')

    def style(ws, widths, money_cols=(), n=0):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws[1]:
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[1].height = 28
        for row in ws.iter_rows(min_row=2, max_row=min(n + 1, ws.max_row)):
            for c in row:
                c.font = body_font
                c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                c.alignment = Alignment(vertical='center',
                                        horizontal='left' if isinstance(c.value, str) else 'right',
                                        wrap_text=isinstance(c.value, str) and len(str(c.value)) > 20)
        for idx in money_cols:
            for i in range(2, n + 2):
                ws.cell(i, idx).number_format = '0.0'

    # 1) 比对结论
    ws0 = wb.active
    ws0.title = '比对结论'
    # 结论只统计"精确差异"笔；候选笔不计入构成（对方未提供逐笔，无法确认）
    det_exact = det[det['线索类型'].str.startswith('精确差异')] if not det.empty else det
    types = (det_exact.groupby('归因类型').agg(笔数=('工号', 'size'), 差额=('本笔差额', 'sum'))
             if not det_exact.empty else pd.DataFrame())

    # 总体判断：总额相同但逐人不同，是最容易误判"已对齐"的情形
    if abs(total_b - total_a) <= 0.011 and len(changed) == 0 and not only_a and not only_b:
        verdict = '逐人逐笔一致，可视为已对齐'
    elif abs(total_b - total_a) <= 0.011 and (len(changed) > 0 or only_a or only_b):
        verdict = f'总额相同（差 {total_b - total_a:+.2f} 元），但仍有 {len(changed)} 人存在差异，正负抵消，不能视为已对齐'
    else:
        verdict = f'总额差 {total_b - total_a:+.2f} 元，{len(changed)} 人存在差异，尚未对齐'

    rows0 = [
        [f'餐费比对结论：{args.a_name} vs {args.b_name}', '', ''],
        ['', '', ''],
        ['项目', args.a_name, args.b_name],
        ['人数', n_a, n_b],
        ['总额（元）', round(total_a, 2), round(total_b, 2)],
        ['总额差（元）', '', round(total_b - total_a, 2)],
        ['仅一方有的人数', len(only_a), len(only_b)],
        ['金额有差异人数', len(changed), ''],
        ['有差异金额合计（元）', '', round(float(changed['差额'].sum()), 2)],
        ['总体判断', verdict, ''],
        ['', '', ''],
        ['结构核验', '', ''],
        [f'{args.a_name}人均口径', A['per_src'], ''],
        [f'{args.b_name}人均口径', B['per_src'], ''],
        [f'逐笔线索笔数', len(det), f'其中精确差异 {len(det_exact)} 笔'],
        ['', '', ''],
        ['差异按归因类型小计（仅精确差异）', '笔数', '差额（元）'],
    ]
    for t, r in types.iterrows():
        rows0.append([t or '未分类', int(r['笔数']), round(float(r['差额']), 2)])
    rows0 += [
        ['', '', ''],
        ['说明', '', ''],
        ['1', f'逐人比对是权威口径；逐笔线索取{args.a_name}明细，{args.b_name}逐笔仅在对方提供明细/差异页时才有'],
        ['2', '归因命中 references/差异归因规则.json 的已归档模式；未命中标"需人工判定"，不硬套'],
        ['3', '对方未提供逐笔时改走逐人净额校验：差额能对上就逐笔列出，对不上的单独一行写明缺口与疑似笔'],
    ]
    for r in rows0:
        ws0.append(r)
    for col, w in zip('ABC', [32, 30, 30]):
        ws0.column_dimensions[col].width = w
    ws0['A1'].font = Font(name='微软雅黑', bold=True, size=13, color='133C9A')
    for row in ws0.iter_rows(min_row=3, max_row=ws0.max_row):
        for c in row:
            c.font = body_font
            c.alignment = Alignment(vertical='center', wrap_text=True)
    for r in (3, 11, 17):
        for c in 'ABC':
            ws0[f'{c}{r}'].fill = PatternFill('solid', fgColor='E1EAFF')
            ws0[f'{c}{r}'].font = Font(name='微软雅黑', bold=True, size=10)
    # 总体判断行突出：未对齐标橙
    ws0.cell(10, 2).fill = PatternFill('solid', fgColor='FCE4D6' if '不能视为已对齐' in verdict or '尚未对齐' in verdict else 'E2EFDA')
    for r in range(4, 10):
        ws0.cell(r, 3).number_format = '0.0'
    ws0.cell(16, 3).number_format = '0'

    # 2) 逐人差异（有差异的人 + 仅一方有的人）
    ws1 = wb.create_sheet('逐人差异')
    cols1 = ['工号', '姓名', f'{args.a_name}金额', f'{args.b_name}金额', '差额', '差异判定']
    ws1.append(cols1)
    allchg = j[(j['差额'].abs() > 0.011) | (j['工号'].isin(only_a | only_b))].copy()
    allchg = allchg.sort_values('差额', ascending=False)
    for _, r in allchg.iterrows():
        if r['工号'] in only_a:
            flag = f'仅{args.a_name}有此人'
        elif r['工号'] in only_b:
            flag = f'仅{args.b_name}有此人'
        else:
            flag = f'{args.b_name}多收' if r['差额'] > 0 else f'{args.a_name}多收'
        ws1.append([r['工号'], r['姓名'], float(r['A金额']), float(r['B金额']), float(r['差额']), flag])
    style(ws1, [14, 12, 14, 14, 10, 22], money_cols=(3, 4, 5), n=len(allchg))

    # 3) 逐笔线索
    ws2 = wb.create_sheet('逐笔线索')
    if not det.empty:
        ws2.append(list(det.columns))
        for _, r in det.iterrows():
            ws2.append(list(r.values))
        style(ws2, [30, 13, 11, 12, 6, 8, 11, 10, 13, 13, 10, 11, 12, 13, 40, 40, 12, 10, 30, 60],
              money_cols=(8, 9, 10, 11), n=len(det))
        # 责任判定含"不符"标浅橙
        for i in range(2, len(det) + 2):
            v = str(ws2.cell(i, 19).value or '')
            if '不符' in v:
                for c in range(1, len(det.columns) + 1):
                    ws2.cell(i, c).fill = PatternFill('solid', fgColor='FCE4D6')
    else:
        ws2.append(['无逐笔线索', 'A 方未提供逐笔明细，只能逐人比对'])
        ws2.column_dimensions['A'].width = 24

    name = f'餐费比对_{args.a_name}_vs_{args.b_name}.xlsx'.replace('/', '-')
    path = os.path.join(args.out, name)
    wb.save(path)

    print(f'{args.a_name}：{n_a} 人 / {total_a:.1f} 元')
    print(f'{args.b_name}：{n_b} 人 / {total_b:.1f} 元')
    print(f'总额差：{total_b - total_a:+.1f} 元；金额有差异 {len(changed)} 人，差异合计 {changed["差额"].sum():+.1f} 元')
    if len(only_a) or len(only_b):
        print(f'仅{args.a_name}有：{sorted(only_a)}；仅{args.b_name}有：{sorted(only_b)}')
    if not types.empty:
        print('差异构成：')
        for t, r in types.iterrows():
            print(f'  {t or "未分类"}：{int(r["笔数"])} 笔 / {r["差额"]:+.1f} 元')
    print(f'输出：{path}')


if __name__ == '__main__':
    main()
