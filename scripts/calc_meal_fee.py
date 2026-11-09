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
    ap.add_argument('--roster', default=None, help='人员名单 xlsx（可选，只核算名单内人员）')
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
    # 加班时段解析 + 归属日期（规则 5.6：归属日期 = 实际打卡结束时间所在日期）
    ot['_st'] = pd.to_datetime(ot.get('实际打卡加班时间'), errors='coerce')
    ot['_en'] = pd.to_datetime(ot.get('实际打卡结束时间'), errors='coerce')
    ot['_mins'] = ot.get('实际加班时长').apply(parse_minutes)
    # 归属日：优先用结束时间日期，缺失回退加班日期
    ot['_attr_date'] = ot.apply(
        lambda r: (r['_en'].date() if pd.notna(r['_en']) else r['加班日期'].date()), axis=1)

    # 按 (工号, 归属日) 预聚合所有加班段（同日多条=分段加班，时长累加）
    OT_MAP = {}
    for _, r in ot.iterrows():
        if pd.isna(r['加班日期']):
            continue
        key = (str(r['员工工号']).strip(), r['_attr_date'])
        seg = (r['_st'], r['_en']) if (pd.notna(r['_st']) and pd.notna(r['_en'])) else None
        e = OT_MAP.setdefault(key, {'segs': [], 'mins': 0})
        e['mins'] += int(r['_mins']) if pd.notna(r['_mins']) else 0
        if seg:
            e['segs'].append(seg)

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

    # ---- 人员名单过滤（可选）：只核算名单内人员 ----
    if args.roster and os.path.exists(args.roster):
        def norm_id(x):
            if pd.isna(x):
                return None
            s = str(x).strip()
            if s.endswith('.0'):
                s = s[:-2]
            if s.isdigit():
                return str(int(s))
            return s.upper()

        xl = pd.ExcelFile(args.roster)
        sh = '合计' if '合计' in xl.sheet_names else xl.sheet_names[0]
        rdf = pd.read_excel(args.roster, sheet_name=sh, dtype=str)
        id_col = None
        for c in rdf.columns:
            if '工号' in str(c):
                id_col = c
                break
        if id_col is None:
            sys.exit(f'名单 {args.roster} 未找到「工号」列，实际列：{list(rdf.columns)}')
        rdf = rdf.dropna(subset=[id_col]).copy()
        rdf['_id'] = rdf[id_col].map(norm_id)
        roster_ids = set(rdf['_id'].dropna())
        meal['_id'] = meal['工号'].map(norm_id)
        n_before, p_before = len(meal), meal['工号'].nunique()
        meal = meal[meal['_id'].isin(roster_ids)].copy()
        print(f'名单过滤：名单 {len(roster_ids)} 人，匹配到就餐记录 {len(meal)} 笔 / '
              f'{meal["工号"].nunique()} 人（过滤前 {n_before} 笔 / {p_before} 人）')
        # 记录名单匹配结果供核对
        matched_ids = set(meal['_id'].unique())
        rdf['_匹配状态'] = rdf['_id'].apply(lambda x: '有就餐记录' if x in matched_ids else '无就餐记录')
        roster_check = rdf.drop(columns=['_id'])
        roster_check.to_csv(os.path.join(args.out, '_名单匹配核对.csv'), index=False, encoding='utf-8-sig')

    # 一天同餐只抵一笔（规则 6.3 / 示例19）：同组内按刷卡时间第一笔参与判定，其余笔一律自付
    meal = meal.sort_values(['工号', '打卡日期', '开餐餐次', '打卡时间']).reset_index(drop=True)
    meal['_first_of_group'] = ~meal.duplicated(subset=['工号', '打卡日期', '开餐餐次'], keep='first')

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

    # ---- 加班覆盖餐别判定（真实时段重叠，含边界）----
    def covered_meals(segs, day):
        """返回该日加班时段覆盖的餐别集合"""
        res = set()
        if not segs:
            return res
        from datetime import timedelta
        for meal, (s_min, e_min) in SLOTS.items():
            ws = datetime.combine(day.date(), time(0, 0)) + timedelta(minutes=s_min)
            we = datetime.combine(day.date(), time(0, 0)) + timedelta(minutes=e_min)
            for st, en in segs:
                if st <= we and en >= ws:
                    res.add(meal)
                    break
        return res

    def early_arrival(segs, day):
        """白班提前到岗：任一加班段开始 ≤ 提前线，且发生在当天上午（<09:00）"""
        cutoff = hm_to_min(cfg['early_cutoff'])
        if cutoff is None:
            return False
        for st, en in segs:
            if st is None:
                continue
            if st.date() == day.date() and (st.hour * 60 + st.minute) <= cutoff:
                return True
        return False

    def late_shift_morning(segs, day):
        """晚班跨天收工次日 05:30-09:00"""
        for st, en in segs:
            if en is None:
                continue
            if en.date() == day.date() and en.hour * 60 + en.minute >= 5 * 60 + 30 \
                    and en.hour * 60 + en.minute <= 9 * 60 and st is not None and st.date() < en.date():
                return True
        return False

    # ---- 逐笔核算 ----
    # 预计算每人每天实际刷卡的餐别集合（周末抵扣须"实际就餐+加班覆盖"同时成立）
    day_meals_map = {}
    for (g, dt), grp in meal.groupby(['工号', '打卡日期']):
        day_meals_map[(g, dt)] = set(grp['开餐餐次'].astype(str).str.strip())

    detail_rows = []
    for _, row in meal.iterrows():
        gh = row['工号']
        name = row.get('姓名', '')
        d = row['打卡日期']
        meal_type = row['开餐餐次']
        amt = float(row['打卡总金额'])
        t_str = row['打卡时间']
        t_min = hm_to_min(t_str)

        # 当日加班（按归属日期聚合，同日分段累加）
        otinfo = OT_MAP.get((gh, d.date()), {'segs': [], 'mins': 0})
        mins = otinfo['mins']
        segs = otinfo['segs']
        covered = covered_meals(segs, d)

        weekend = is_weekend(d)

        # 工作日午餐 4 小时门槛：班次时长 <4h 不豁免
        shift_len = 0
        for (ss, se) in get_shift(gh):
            if se > ss:
                shift_len = max(shift_len, se - ss)

        # 判断工作餐：仅工作日；刷卡时间落在班次时段内 → 免费
        # ① 早餐不适用工作餐（规则 5.4 优先：早餐仅提前到岗/跨天收工两种情况可抵）
        # ② 班次结束时刻取严格小于：17:15 下班后刷卡（如 17:15:03）不算工作餐
        # ③ 工作日午餐须班次≥4h 才免费（规则第四章备注①）
        is_work_meal = False
        if not weekend and t_min is not None and meal_type != '早餐':
            for (ss, se) in get_shift(gh):
                if ss <= t_min < se:
                    is_work_meal = True
                    break
        if is_work_meal and meal_type == '午餐' and not weekend and shift_len < 4 * 60:
            is_work_meal = False

        deduct = 0.0
        reason = ''

        if not row['_first_of_group']:
            # 同日同餐第 2 笔及以后：不再抵扣，全额自付（规则 6.3 / 示例19）
            reason = '同日同餐别第2笔及以后（一天同餐只抵一笔），全额自付'
        elif is_work_meal:
            deduct = amt
            reason = '工作餐（班次时段内），免费'
        elif meal_type == '早餐':
            # 早餐特殊规则（规则 5.4）：仅①白班提前到岗 ②晚班跨天收工次日
            if early_arrival(segs, d):
                deduct = min(amt, cfg['breakfast_amt'])
                reason = f'白班提前到岗（开始≤{cfg["early_cutoff"]}），抵早餐'
            elif late_shift_morning(segs, d):
                deduct = min(amt, cfg['breakfast_amt'])
                reason = '晚班跨天收工次日（≥05:30），抵早餐'
            elif not segs:
                reason = '早餐：无加班记录，不抵'
            else:
                reason = f'早餐不达标（加班{mins}min，开始{segs[0][0].strftime("%m-%d %H:%M") if segs[0][0] else "-"}），自付'
        elif not weekend:
            # 工作日
            if meal_type == '午餐':
                if shift_len >= 4 * 60:
                    deduct = amt
                    reason = '工作日午餐（班次≥4h），免费'
                else:
                    deduct = 0
                    reason = f'工作日午餐但班次{shift_len}min<4h，不豁免扣5元'
            elif meal_type == '晚餐':
                # 须加班≥1h 且 加班时段覆盖晚餐时段（规则 5.3 / 示例20）
                if mins >= 60 and '晚餐' in covered:
                    deduct = min(amt, cfg['dinner_amt'])
                    reason = f'工作日加班{mins}min≥1h，覆盖晚餐，抵5元'
                elif mins >= 60:
                    reason = f'工作日加班{mins}min≥1h但未覆盖晚餐时段，不抵'
                else:
                    reason = f'工作日加班{mins}min<1h，晚餐不抵'
            else:
                reason = '工作日非午餐/晚餐餐别，自付'
        else:
            # 周末/节假日：先按时长定可抵餐次数，再按 晚餐>午餐>早餐 从「实际覆盖的餐别」中抵（未覆盖作废）
            if mins > 300:
                n_meal = 2
            elif mins >= 180:
                n_meal = 1
            else:
                n_meal = 0
            # 周末早餐不参与（仅提前到岗/跨天可抵，本分支不处理早餐）
            # 只在「加班覆盖 + 当天实际刷卡」的餐别中抵扣：避免只吃了午餐却把额度作废给晚餐
            dm = day_meals_map.get((gh, d), set())
            avail = [m for m in ('晚餐', '午餐') if m in covered and m in dm]
            target = None
            for m in avail:
                if m == meal_type:
                    target = m
                    break
            if n_meal > 0 and meal_type in avail:
                # 该餐是优先序列中的第几顺位，是否在可抵名额内
                rank = avail.index(meal_type) + 1
                if rank <= n_meal:
                    deduct = min(amt, cfg['dinner_amt'])
                    reason = f'周末加班{mins}min，可抵{n_meal}餐，覆盖{meal_type}，抵5元'
                else:
                    reason = f'周末加班{mins}min可抵{n_meal}餐，按晚餐优先，{meal_type}不抵'
            elif meal_type == '早餐':
                if early_arrival(segs, d):
                    deduct = min(amt, cfg['breakfast_amt'])
                    reason = '周末白班提前到岗，抵早餐'
                elif late_shift_morning(segs, d):
                    deduct = min(amt, cfg['breakfast_amt'])
                    reason = '周末晚班跨天收工次日，抵早餐'
                else:
                    reason = f'周末加班{mins}min，早餐不抵'
            else:
                reason = f'周末加班{mins}min<3h或无加班/未覆盖{meal_type}，{meal_type}自付'

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
