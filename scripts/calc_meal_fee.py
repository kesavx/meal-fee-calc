#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
餐费核算脚本（meal-fee-calc）

按规则基准.md + regions.csv 配置，用加班记录+食堂打卡+班次表核算每位员工应扣餐费。

用法：
    python3 calc_meal_fee.py \
      --ot <加班记录.xlsx> \
      --meal <食堂打卡.xlsx> \
      [--shift <班次表.xlsx>] \
      --region 广东 \
      --out <输出目录>

输入列要求：
  加班记录：员工工号、加班日期、实际打卡加班时间、实际打卡结束时间、实际加班时长、单据类型
  食堂打卡：工号、姓名、打卡日期、打卡时间、开餐餐次、打卡总金额
  班次表(可选)：含员工工号/工号 与班次上下班时间（列名可配置）

输出：<out>/逐笔明细.csv 与 <out>/按人汇总.csv
"""
import argparse
import csv
import os
import re
import sys
from datetime import datetime, time

import pandas as pd

# ---------------- 区域配置加载 ----------------
REGIONS_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'regions.csv')

def load_regions(path=REGIONS_CSV):
    df = pd.read_csv(path, dtype=str)
    regions = {}
    for _, r in df.iterrows():
        regions[r['region']] = {
            'breakfast_amt': float(r['breakfast_amt']),
            'lunch_amt': float(r['lunch_amt']),
            'dinner_amt': float(r['dinner_amt']),
            'shift_start': r['default_shift_start'],
            'shift_end': r['default_shift_end'],
            'early_cutoff': r['early_arrival_cutoff'],
            'note': r['note'] if pd.notna(r.get('note')) else '',
        }
    return regions

# ---------------- 时间工具 ----------------
def parse_minutes(s):
    """从 '120分钟' / '2小时' / 120 / nan 解析出分钟数"""
    if pd.isna(s):
        return 0
    s = str(s).strip()
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if not m:
        return 0
    return int(float(m.group(1)))

def parse_hm(s):
    """解析 '08:30' / '8:30' / '2026-09-01 08:30:00' 为 (h, m)"""
    if pd.isna(s):
        return None
    s = str(s).strip()
    m = re.search(r'(\d{1,2}):(\d{2})', s)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))

def hm_to_min(s):
    p = parse_hm(s)
    return p[0] * 60 + p[1] if p else None

def in_slot(t_min, start_hm, end_hm):
    """刷卡时间是否落在餐别时段内（含边界）"""
    if t_min is None:
        return False
    return start_hm <= t_min <= end_hm

# 餐别时段（规则基准.md 第二章）
SLOTS = {
    '早餐': (5 * 60 + 30, 9 * 60),       # 05:30-09:00
    '午餐': (10 * 60 + 30, 14 * 60 + 30),  # 10:30-14:30
    '晚餐': (15 * 60 + 30, 20 * 60 + 30),  # 15:30-20:30
}

# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser(description='餐费核算')
    ap.add_argument('--ot', required=True, help='加班记录 xlsx')
    ap.add_argument('--meal', required=True, help='食堂打卡 xlsx')
    ap.add_argument('--shift', default=None, help='班次表 xlsx（可选）')
    ap.add_argument('--region', required=True, help='区域名，需在 regions.csv 中存在')
    ap.add_argument('--out', required=True, help='输出目录')
    ap.add_argument('--adjust-workdays', default='2026-09-20',
                    help='调休上班日（逗号分隔，按工作日规则），默认 2026-09-20')
    args = ap.parse_args()

    regions = load_regions()
    if args.region not in regions:
        sys.exit(f'区域 {args.region} 不在 regions.csv 中，可选：{list(regions.keys())}')
    cfg = regions[args.region]

    os.makedirs(args.out, exist_ok=True)

    # ---- 读数据 ----
    # 加班记录可能带一行英文系统表头（header=1 才是中文表头），自动探测
    def smart_read_ot(path):
        raw = pd.read_excel(path, header=None, nrows=2)
        row0 = ' '.join([str(x) for x in raw.iloc[0]])
        row1 = ' '.join([str(x) for x in raw.iloc[1]])
        if '加班日期' in row1 and '员工工号' in row1:
            return pd.read_excel(path, header=1)
        if '加班日期' in row0 or 'OverTimeDate' in row0:
            return pd.read_excel(path, header=0)
        return pd.read_excel(path)

    ot = smart_read_ot(args.ot)
    meal = pd.read_excel(args.meal)

    # 加班记录列名归一（兼容中英文表头）
    col_map_ot = {}
    for c in ot.columns:
        cs = str(c).strip()
        if '员工工号' in cs or cs in ('StaffId-StaffCode',):
            col_map_ot['员工工号'] = c
        elif '加班日期' in cs or cs == 'OverTimeDate':
            col_map_ot['加班日期'] = c
        elif '实际打卡加班时间' in cs or cs == 'ActualStartDateTime':
            col_map_ot['实际打卡加班时间'] = c
        elif '实际打卡结束时间' in cs or cs == 'ActualStopDateTime':
            col_map_ot['实际打卡结束时间'] = c
        elif '实际加班时长' in cs or cs == 'ActualDurationIncludeUnit':
            col_map_ot['实际加班时长'] = c
        elif '单据类型' in cs or cs == 'DocumentType':
            col_map_ot['单据类型'] = c
    ot = ot.rename(columns={v: k for k, v in col_map_ot.items()})
    ot['员工工号'] = ot['员工工号'].astype(str).str.strip()
    # 排除撤销
    if '单据类型' in ot.columns:
        ot = ot[ot['单据类型'].astype(str).str.strip() != '撤销'].copy()
    ot['加班日期'] = pd.to_datetime(ot['加班日期'])

    # 食堂打卡列名归一（避免"开餐餐次"/"餐次"等近似列冲突）
    meal_cols = {str(c).strip(): c for c in meal.columns}
    def pick(prefer, fallback=None):
        for p in prefer:
            if p in meal_cols:
                return meal_cols[p]
        for f in (fallback or []):
            for k, v in meal_cols.items():
                if f in k:
                    return v
        return None

    col_map_meal = {
        '工号': pick(['工号', '员工工号']),
        '姓名': pick(['姓名']),
        '打卡日期': pick(['打卡日期']),
        '打卡时间': pick(['打卡时间']),
        '开餐餐次': pick(['开餐餐次'], ['餐次']),
        '打卡总金额': pick(['打卡总金额', '实际员工支出金额'], ['金额']),
    }
    col_map_meal = {k: v for k, v in col_map_meal.items() if v}
    missing = [k for k in ['工号', '打卡日期', '打卡时间', '开餐餐次', '打卡总金额'] if k not in col_map_meal]
    if missing:
        sys.exit(f'食堂打卡缺少必需列：{missing}，实际列：{list(meal.columns)}')
    meal = meal.rename(columns={v: k for k, v in col_map_meal.items()})
    meal['工号'] = meal['工号'].astype(str).str.strip()
    meal['打卡日期'] = pd.to_datetime(meal['打卡日期'])
    meal['打卡时间'] = meal['打卡时间'].astype(str)
    meal['开餐餐次'] = meal['开餐餐次'].astype(str).str.strip()
    meal['打卡总金额'] = pd.to_numeric(meal['打卡总金额'], errors='coerce').fillna(0)

    # 去重：同工号+日期+餐别只留一笔
    meal = meal.drop_duplicates(subset=['工号', '打卡日期', '开餐餐次'], keep='first').copy()

    # ---- 班次表（可选）----
    shift_map = {}  # 工号 -> (start_hm, end_hm) 分钟
    if args.shift and os.path.exists(args.shift):
        sh = pd.read_excel(args.shift)
        id_col = None
        start_col = None
        end_col = None
        for c in sh.columns:
            cs = str(c).strip()
            if '员工工号' in cs or cs == '工号':
                id_col = c
            elif '上班时间' in cs or '开始' in cs:
                start_col = c
            elif '下班时间' in cs or '结束' in cs:
                end_col = c
        if id_col and start_col and end_col:
            for _, r in sh.iterrows():
                try:
                    sid = str(r[id_col]).strip()
                    s = hm_to_min(str(r[start_col]))
                    e = hm_to_min(str(r[end_col]))
                    if sid and s is not None and e is not None:
                        shift_map.setdefault(sid, []).append((s, e))
                except Exception:
                    continue

    def get_shift(gh):
        """返回该员工当天班次列表（可能多段），无则用区域默认班次"""
        if gh in shift_map and shift_map[gh]:
            return shift_map[gh]
        s = hm_to_min(cfg['shift_start'])
        e = hm_to_min(cfg['shift_end'])
        return [(s, e)] if s is not None and e is not None else []

    # ---- 调休日 ----
    adjust_days = set()
    for d in str(args.adjust_workdays).split(','):
        d = d.strip()
        if d:
            adjust_days.add(pd.Timestamp(d).date())

    def is_weekend(d):
        return d.weekday() >= 5 and d.date() not in adjust_days

    # ---- 逐笔核算 ----
    detail_rows = []
    for _, row in meal.iterrows():
        gh = row['工号']
        name = row.get('姓名', '')
        d = row['打卡日期']
        meal_type = row['开餐餐次']
        amt = float(row['打卡总金额'])
        t_str = row['打卡时间']
        t_min = hm_to_min(t_str)

        # 当日加班
        day_ot = ot[(ot['员工工号'] == gh) & (ot['加班日期'] == d)]
        mins = parse_minutes(day_ot['实际加班时长'].iloc[0]) if len(day_ot) > 0 else 0

        weekend = is_weekend(d)

        # 判断工作餐：仅工作日；刷卡时间落在班次时段内 → 免费
        is_work_meal = False
        if not weekend and t_min is not None:
            for (ss, se) in get_shift(gh):
                if ss <= t_min <= se:
                    is_work_meal = True
                    break

        deduct = 0.0
        reason = ''

        if is_work_meal:
            deduct = amt
            reason = '工作餐（班次时段内），免费'
        elif meal_type == '早餐':
            # 早餐特殊规则（规则基准 5.4）
            if not day_ot.empty:
                start_str = str(day_ot['实际打卡加班时间'].iloc[0])
                end_str = str(day_ot['实际打卡结束时间'].iloc[0])
                start_dt, end_dt = None, None
                for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
                    for s, attr in ((start_str, 'start_dt'), (end_str, 'end_dt')):
                        try:
                            parsed = datetime.strptime(s.strip(), fmt)
                            if attr == 'start_dt':
                                start_dt = parsed
                            else:
                                end_dt = parsed
                        except ValueError:
                            continue
                start_min = start_dt.hour * 60 + start_dt.minute if start_dt else hm_to_min(start_str)
                end_min = end_dt.hour * 60 + end_dt.minute if end_dt else hm_to_min(end_str)
                crosses_day = bool(start_dt and end_dt and end_dt.date() > start_dt.date())
                # ① 白班提前到岗：开始时间在凌晨/上午（<09:00）且 ≤ 提前线
                cutoff = hm_to_min(cfg['early_cutoff'])
                if start_min is not None and start_min < 9 * 60 and cutoff is not None and start_min <= cutoff:
                    deduct = min(amt, cfg['breakfast_amt'])
                    reason = f'白班提前到岗（开始≤{cfg["early_cutoff"]}），抵早餐'
                # ② 晚班后次日早上：跨天且收工落在早餐时段（05:30-09:00）
                elif crosses_day and end_min is not None and 5 * 60 + 30 <= end_min <= 9 * 60:
                    deduct = min(amt, cfg['breakfast_amt'])
                    reason = f'晚班跨天收工次日{end_dt.strftime("%H:%M")}，抵早餐'
                else:
                    reason = f'早餐不达标（开始{start_str}，结束{end_str}），自付'
            else:
                reason = '早餐：无加班记录，不抵'
        elif not weekend:
            # 工作日
            if meal_type == '午餐':
                # 4小时门槛：班次时长≥4h免费，否则扣5元
                shift_len = 0
                for (ss, se) in get_shift(gh):
                    if se > ss:
                        shift_len = max(shift_len, se - ss)
                if shift_len >= 4 * 60:
                    deduct = amt
                    reason = '工作日午餐（班次≥4h），免费'
                else:
                    deduct = 0
                    reason = f'工作日午餐但班次时长{shift_len}min<4h，不豁免扣5元'
            elif meal_type == '晚餐':
                if mins >= 60:
                    deduct = min(amt, cfg['dinner_amt'])
                    reason = f'工作日加班{mins}min≥1h，抵晚餐'
                else:
                    reason = f'工作日加班{mins}min<1h，晚餐不抵'
            else:
                reason = '工作日非午餐/晚餐餐别，自付'
        else:
            # 周末/节假日
            if mins > 300:
                if meal_type in ('午餐', '晚餐'):
                    deduct = min(amt, cfg['dinner_amt'])
                    reason = f'周末加班{mins}min>5h，抵{meal_type}'
                else:
                    reason = '周末加班>5h，早餐不抵'
            elif mins >= 180:
                if meal_type == '晚餐':
                    deduct = min(amt, cfg['dinner_amt'])
                    reason = f'周末加班{mins}min(3-5h)，抵晚餐'
                else:
                    reason = f'周末加班{mins}min(3-5h)，优先抵晚餐，{meal_type}不抵'
            else:
                reason = f'周末加班{mins}min<3h或无加班，{meal_type}自付'

        self_pay = round(max(0, amt - deduct), 2)
        detail_rows.append({
            '工号': gh, '姓名': name,
            '日期': d.strftime('%Y-%m-%d'),
            '星期': '一二三四五六日'[d.weekday()],
            '餐别': meal_type, '刷卡金额': amt,
            '抵扣金额': deduct, '自付金额': self_pay,
            '加班时长(分)': mins,
            '判定': reason,
        })

    detail_df = pd.DataFrame(detail_rows)

    # ---- 按人汇总 ----
    summary = detail_df.groupby(['工号', '姓名'], as_index=False)['自付金额'].sum()
    summary = summary.rename(columns={'自付金额': '应扣餐费'}).sort_values('应扣餐费', ascending=False)

    # ---- 输出 ----
    detail_path = os.path.join(args.out, '逐笔明细.csv')
    summary_path = os.path.join(args.out, '按人汇总.csv')
    detail_df.to_csv(detail_path, index=False, encoding='utf-8-sig')
    summary.to_csv(summary_path, index=False, encoding='utf-8-sig')

    total = detail_df['自付金额'].sum()
    print(f'区域：{args.region}')
    print(f'人数：{len(summary)}，明细笔数：{len(detail_df)}')
    print(f'应扣总额：{total:.2f} 元')
    print(f'输出：{detail_path}')
    print(f'输出：{summary_path}')

if __name__ == '__main__':
    main()
