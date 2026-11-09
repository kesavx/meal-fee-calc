# meal-fee-calc

员工餐费核算技能：一份规则仓库 + 两个工作流（核算 / 两版核对分析）。

## 它能做什么

| 能力 | 说明 |
|---|---|
| 规则仓库 | 只有一份现行规则（`references/规则基准.md`），每次核算都取最新版；修订直接覆盖并追加变更记录，不保留按月分版的旧文件 |
| 核算 | 用「加班记录 + 食堂打卡 + 班次表」算出每位员工应扣餐费，输出逐笔明细（哪天、哪餐、为什么抵/为什么自付）和按人汇总 |
| 核对分析 | 输入两份餐费结果，逐笔定位差异、按已归档模式归因、判定哪一方与现行规则不符；总额相同但逐人不同时会明确标出「正负抵消，不能视为已对齐」 |

## 安装

把整个 `meal-fee-calc` 目录复制到客户端的 user skills 目录下，重启客户端即可。

该目录位于 `.../agent_mode/workspace/.user_skills/`，具体父路径随操作系统与运行环境不同，以本机已有的技能目录为准，不要自己新建一个位置。

```
<已存在的 workspace>/.user_skills/meal-fee-calc/
├── SKILL.md
├── references/
└── scripts/
```

装好后无需任何配置，提到「算餐费」「核对两份餐费」「改餐费规则」时会自动命中。

## 依赖

- Python 3
- `pandas`、`openpyxl`（读 xlsx 与写比对结果）

```bash
pip3 install pandas openpyxl
```

## 快速开始

核算：

```bash
python3 scripts/calc_meal_fee.py \
  --ot 加班记录.xlsx --meal 食堂打卡.xlsx \
  --region 广东 \
  --roster 试点人员名单.xlsx \
  --adjust-workdays 2026-09-20,2026-10-10 \
  --out ./out
```

比对两份结果：

```bash
python3 scripts/compare_meal_fee.py \
  --a ./out/逐笔明细.csv --b 他人结果.xlsx \
  --ot 加班记录.xlsx --shift-end 17:15 \
  --out ./cmp --a-name 本方 --b-name 对方
```

## 文件说明

| 文件 | 作用 |
|---|---|
| `SKILL.md` | 技能入口：开工顺序、三个工作流、易错点 |
| `references/规则基准.md` | 唯一现行规则 + 变更记录 |
| `references/区域配置表.md` | 各区域餐标、默认班次、特殊说明 |
| `references/差异归因规则.json` | 比对用的已归档差异模式 |
| `scripts/regions.csv` | 区域配置的机器可读版 |
| `scripts/calc_meal_fee.py` | 核算 |
| `scripts/compare_meal_fee.py` | 两版比对、归因、责任判定 |

## 维护约定

1. 改规则只改 `references/规则基准.md`，并在文末「变更记录」追加一行（日期 / 变更内容 / 触发来源）。
2. 新规则只对下一次核算生效，不追溯已交付结果。
3. 区域差异只进配置表和 `regions.csv`；新增区域 = 各加一行，规则基准与脚本都不用动。
4. 改判定逻辑要同步改脚本并重跑验证，不要只改文档。
5. 不要把内部餐标、人员名单、核算结果提交到本仓库。

> 本仓库为私有，内含公司内部核算口径，请勿公开或外传。
