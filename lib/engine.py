#!/usr/bin/env python3
"""台区线损与台区异常 —— 确定性计算内核。

设计约束：
1. 所有数值来自工作簿本身或调用参数，不做估计、不填默认阈值。
2. 缺失字段 / 缺失参数一律进入 gaps，交给上层写成"待补充"。
3. 输出固定 JSON（含 markdown 摘要），供 Host 插件工具直接渲染。

用法： python3 engine.py <scan|diagnose> '<json args>'
"""
import json
import os
import re
import sys

try:
    import openpyxl
except Exception as exc:  # pragma: no cover
    print(json.dumps({"ok": False, "error": f"openpyxl 不可用: {exc}"}, ensure_ascii=False))
    sys.exit(0)


ROLE_KEYWORDS = {
    "tgId": [["台区编号"], ["台区id"], ["台区ID"], ["tgid"]],
    "tgName": [["台区名称"], ["台区简称"], ["简称"], ["台区"]],
    "date": [["数据日期"], ["日期"], ["统计日期"]],
    "terminal": [["终端地址码"], ["终端编号"], ["终端资产编号"], ["终端"]],
    "meter": [["电表条码"], ["电能表条码"], ["电能表资产编号"], ["电表号"], ["表号"]],
    "mp": [["测量点号"]],
    "customer": [["客户编号"], ["户号"], ["客户名称"]],
    "ratio": [["综合倍率"], ["倍率"]],
    "tgFlag": [["Is Tg Meter"], ["总表标识"]],
    "supply": [["供电量"], ["供入电量"], ["总表电量"]],
    "sale": [["售电量"], ["售出电量"], ["户表电量"]],
    "lossRate": [["线损率"]],
    "lossEnergy": [["线损电量"]],
    "station": [["所在配电室"], ["配电室"]],
    "crossCount": [["跨台区数"], ["覆盖台区数"]],
    "totalTerminal": [["本终端对应的台区总表"], ["台区总表所在终端"]],
}

# 派生列不参与角色识别：这些列是上一次分析的结论，不是原始档案字段
BLOCK_TOKENS = ("应属", "档案", "现场", "核实", "迁移", "单元主", "多数", "修正", "方案", "对比", "偏差", "安装", "厂商")

DAY_COL = re.compile(r"^\d{1,2}[-/月]\d{1,2}日?$")


def to_num(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def norm(value):
    if value is None:
        return ""
    return str(value).strip()


def read_workbook(path, sheet_name=None):
    """返回 (sheet, header_row_no, columns, rows, sheet_list, source_note)。"""
    ext = os.path.splitext(path)[1].lower()
    if not os.path.exists(path):
        raise FileNotFoundError(f"文件不存在: {path}")
    if ext in (".xls", ".xlsb"):
        # 老格式交给 pandas（若可用），失败则要求转换
        try:
            import pandas as pd

            raw = pd.read_excel(path, sheet_name=0 if sheet_name is None else sheet_name, header=None)
            grid = raw.values.tolist()
            sheet_list = [{"name": str(0 if sheet_name is None else sheet_name), "rows": len(grid)}]
            note = "通过 pandas 读取 .xls（表头行按内容探测）"
            return _from_grid(str(sheet_name or 0), grid, sheet_list, note)
        except Exception as exc:
            raise RuntimeError(
                f"读取 {ext} 失败（{exc}）。请先用 LibreOffice 转成 .xlsx 再诊断。"
            )
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet_list = [{"name": ws.title, "rows": ws.max_row, "cols": ws.max_column} for ws in wb.worksheets]
    ws = None
    if sheet_name:
        ws = wb[sheet_name]
    else:
        best = None
        for cand in wb.worksheets:
            grid = _grid_of(cand, limit=30)
            hit = _header_score(grid)
            if best is None or hit > best[0]:
                best = (hit, cand)
        ws = best[1] if best else wb.worksheets[0]
    grid = _grid_of(ws, limit=None)
    wb.close()
    return _from_grid(ws.title, grid, sheet_list, "openpyxl 直接读取")


def _grid_of(ws, limit=None):
    rows = []
    for idx, row in enumerate(ws.iter_rows(values_only=True)):
        rows.append(list(row))
        if limit is not None and idx + 1 >= limit:
            break
    return rows


def _all_keywords():
    out = []
    for group in ROLE_KEYWORDS.values():
        for aliases in group:
            for alias in aliases:
                out.append(alias.lower())
    return out


def _header_score(grid):
    keys = _all_keywords()
    best = 0
    for row in grid:
        cells = [norm(c).lower() for c in row if norm(c)]
        if not cells:
            continue
        hit = sum(1 for c in cells if any(k in c for k in keys))
        best = max(best, hit)
    return best


def _from_grid(sheet, grid, sheet_list, note):
    if not grid:
        raise RuntimeError("工作表为空")
    keys = _all_keywords()
    header_row = 0
    best = -1
    for idx, row in enumerate(grid[:25]):
        cells = [norm(c).lower() for c in row if norm(c)]
        if not cells:
            continue
        hit = sum(1 for c in cells if any(k in c for k in keys))
        if hit > best:
            best, header_row = hit, idx
    if best < 3:
        header_row = 0
    columns = [norm(c) for c in grid[header_row]]
    # 去掉尾部空列名
    while columns and not columns[-1]:
        columns.pop()
    rows = []
    for raw in grid[header_row + 1:]:
        if all(norm(c) == "" for c in raw):
            continue
        item = {}
        for i, col in enumerate(columns):
            if not col:
                continue
            item[col] = raw[i] if i < len(raw) else None
        rows.append(item)
    return {
        "sheet": sheet,
        "headerRow": header_row + 1,
        "columns": columns,
        "rows": rows,
        "sheetList": sheet_list,
        "sourceNote": note,
    }


def guess_roles(columns):
    roles = {}
    used = set()
    lowered = [(c, c.lower()) for c in columns]
    for role, groups in ROLE_KEYWORDS.items():
        for aliases in groups:
            found = None
            for alias in aliases:
                for col, low in lowered:
                    if col in used or any(tok in col for tok in BLOCK_TOKENS):
                        continue
                    if low == alias.lower():
                        found = col
                        break
                if found:
                    break
            if found is None:
                for alias in aliases:
                    for col, low in lowered:
                        if col in used or any(tok in col for tok in BLOCK_TOKENS):
                            continue
                        if alias.lower() in low:
                            found = col
                            break
                    if found:
                        break
            if found:
                roles[role] = found
                used.add(found)
                break
    return roles


def energy_columns(columns, roles):
    """把可作为电量的列收集起来：Pap E 系列、日列、单列电量。"""
    picked = []
    for col in columns:
        low = col.lower()
        if col in roles.values():
            continue
        if re.fullmatch(r"pap e(\d+)?", low) or re.fullmatch(r"pap_e(\d+)?", low):
            picked.append(col)
        elif DAY_COL.match(col):
            picked.append(col)
        elif low in ("电量", "日电量", "电量kwh", "期间电量", "期间电量kwh", "七日电量", "8日电量", "七日电kwh"):
            picked.append(col)
    if not picked:
        for col in columns:
            low = col.lower()
            if "电量" in low and "线损" not in low and "供电" not in low and "售电" not in low:
                picked.append(col)
                break
    return picked


def sum_energy(row, cols):
    total = 0.0
    seen = False
    for col in cols:
        val = to_num(row.get(col))
        if val is not None:
            total += val
            seen = True
    return total if seen else None


def _date_range(rows, date_col):
    if not date_col:
        return None
    days = [norm(r.get(date_col)) for r in rows if norm(r.get(date_col))]
    if not days:
        return None
    return {"min": min(days), "max": max(days), "distinct": len(set(days))}


def _dedupe(rows, date_col, meter_col, mp_col):
    groups = {}
    for row in rows:
        meter = norm(row.get(meter_col)) if meter_col else ""
        day = norm(row.get(date_col)) if date_col else ""
        if not meter:
            continue
        groups.setdefault((day, meter), []).append(row)
    dup_keys = {k: v for k, v in groups.items() if len(v) > 1}
    kept = []
    dropped = 0
    for key, items in groups.items():
        if len(items) == 1:
            kept.append(items[0])
            continue
        chosen = None
        if mp_col:
            for item in items:
                if norm(item.get(mp_col)):
                    chosen = item
                    break
        chosen = chosen or items[0]
        kept.append(chosen)
        dropped += len(items) - 1
    return kept, dup_keys, dropped


def scan(args):
    path = args["file"]
    data = read_workbook(path, args.get("sheet"))
    columns = data["columns"]
    roles = guess_roles(columns)
    dup = None
    if roles.get("date") and roles.get("meter"):
        _, dup_keys, dropped = _dedupe(data["rows"], roles["date"], roles["meter"], roles.get("mp"))
        dup = {"groups": len(dup_keys), "rowsDropped": dropped, "key": f'{roles["date"]} + {roles["meter"]}'}
    elif roles.get("meter"):
        _, dup_keys, dropped = _dedupe(data["rows"], None, roles["meter"], roles.get("mp"))
        dup = {"groups": len(dup_keys), "rowsDropped": dropped, "key": f'{roles["meter"]}（无日期列）'}
    result = {
        "ok": True,
        "file": os.path.abspath(path),
        "sheet": data["sheet"],
        "sheetList": data["sheetList"],
        "headerRow": data["headerRow"],
        "rowCount": len(data["rows"]),
        "columns": columns,
        "mapping": roles,
        "energyColumns": energy_columns(columns, roles),
        "dateRange": _date_range(data["rows"], roles.get("date")),
        "dedupe": dup,
        "sourceNote": data["sourceNote"],
        "gaps": [],
    }
    for role in ("tgId", "terminal", "meter"):
        if not roles.get(role):
            result["gaps"].append(f"未识别到 {role} 列，相关规则无法执行")
    result["markdown"] = render_scan(result)
    return result


def render_scan(res):
    lines = [f"# 表结构扫描 · {os.path.basename(res['file'])}", ""]
    lines.append(f"- 工作表：`{res['sheet']}`（共 {len(res['sheetList'])} 张）")
    lines.append(f"- 表头行：第 {res['headerRow']} 行（按关键词命中数探测，表头常不在第 1 行）")
    lines.append(f"- 数据行：{res['rowCount']} 行；列数：{len(res['columns'])}")
    if res["dateRange"]:
        d = res["dateRange"]
        lines.append(f"- 日期范围：{d['min']} ~ {d['max']}，去重后 {d['distinct']} 个周期")
    lines.append("")
    lines.append("## 字段角色")
    lines.append("")
    lines.append("| 角色 | 列名 |")
    lines.append("|---|---|")
    label = {
        "tgId": "台区标识",
        "tgName": "台区名称",
        "date": "周期",
        "terminal": "终端",
        "meter": "电表条码",
        "mp": "测量点号",
        "customer": "客户",
        "ratio": "综合倍率",
        "tgFlag": "总表标识",
        "supply": "供电量",
        "sale": "售电量",
        "lossRate": "线损率",
        "lossEnergy": "线损电量",
        "station": "配电室",
    }
    for role, col in res["mapping"].items():
        lines.append(f"| {label.get(role, role)} | `{col}` |")
    if res["energyColumns"]:
        lines.append(f"| 电量列（{len(res['energyColumns'])} 个） | {', '.join('`%s`' % c for c in res['energyColumns'][:8])}{' …' if len(res['energyColumns']) > 8 else ''} |")
    lines.append("")
    if res["dedupe"]:
        d = res["dedupe"]
        lines.append("## 重复行（去重键）")
        lines.append("")
        lines.append(f"- 键：{d['key']}")
        lines.append(f"- 重复组：**{d['groups']}** 组，可精简 **{d['rowsDropped']}** 行")
        lines.append("- 保留规则：同一键内优先保留 `测量点号` 非空的那条")
        lines.append("")
    if res["gaps"]:
        lines.append("## 缺口")
        lines.append("")
        for g in res["gaps"]:
            lines.append(f"- {g}")
        lines.append("")
    lines.append(f"> 来源：{res['sourceNote']}。本表仅含文件内事实，未做任何阈值判断。")
    return "\n".join(lines)


def diagnose(args):
    path = args["file"]
    params = args.get("params") or {}
    data = read_workbook(path, args.get("sheet"))
    columns = data["columns"]
    roles = guess_roles(columns)
    overrides = args.get("mapping") or {}
    mapping_source = "自动识别"
    if overrides:
        for role, col in overrides.items():
            if col in columns:
                roles[role] = col
        mapping_source = "自动识别 + 调用方覆盖"

    gaps = []
    energies = energy_columns(columns, roles)
    rows, dup_keys, dropped = _dedupe(
        data["rows"], roles.get("date"), roles.get("meter"), roles.get("mp")
    ) if roles.get("meter") else (data["rows"], {}, 0)

    by_area_supply = {}
    by_area_sale = {}
    area_days = {}
    detail = {}

    has_supply = bool(roles.get("supply"))
    has_sale = bool(roles.get("sale"))
    tg_flag = roles.get("tgFlag")

    def area_key(row):
        return norm(row.get(roles.get("tgId"))) or norm(row.get(roles.get("tgName"))) or "（未知台区）"

    # 逐行归类：总表 / 户表
    for row in rows:
        area = area_key(row)
        day = norm(row.get(roles.get("date"))) if roles.get("date") else "全期"
        info = detail.setdefault(area, {"terminals": set(), "totalTerminals": set(), "users": 0, "userRows": 0})
        terminal = norm(row.get(roles.get("terminal"))) if roles.get("terminal") else ""
        if terminal:
            info["terminals"].add(terminal)
        is_total = None
        if tg_flag:
            flag = norm(row.get(tg_flag))
            is_total = flag in ("是", "Y", "y", "1", "True", "true")
        if is_total:
            if terminal:
                info["totalTerminals"].add(terminal)
        else:
            info["users"] += 1

        supply_val = to_num(row.get(roles["supply"])) if has_supply else None
        sale_val = to_num(row.get(roles["sale"])) if has_sale else None
        if supply_val is not None:
            by_area_supply[area] = by_area_supply.get(area, 0.0) + supply_val
            area_days.setdefault(area, {}).setdefault(day, {"supply": 0.0, "sale": 0.0})
            area_days[area][day]["supply"] += supply_val
        if sale_val is not None:
            by_area_sale[area] = by_area_sale.get(area, 0.0) + sale_val
            area_days.setdefault(area, {}).setdefault(day, {"supply": 0.0, "sale": 0.0})
            area_days[area][day]["sale"] += sale_val
        if supply_val is None and sale_val is None and energies and tg_flag:
            energy = sum_energy(row, energies)
            if energy is not None:
                bucket = "supply" if is_total else "sale"
                if bucket == "supply":
                    by_area_supply[area] = by_area_supply.get(area, 0.0) + energy
                else:
                    by_area_sale[area] = by_area_sale.get(area, 0.0) + energy

    if not has_supply and not has_sale and not (energies and tg_flag):
        gaps.append("未找到供电量/售电量列，也没有可用的电量列与总表标识，B 类线损计算无法执行")

    areas = []
    for area, supply in sorted(by_area_supply.items()):
        sale = by_area_sale.get(area, 0.0)
        loss = supply - sale
        rate = (loss / supply * 100.0) if supply else None
        days = area_days.get(area, {})
        neg_days = sum(1 for v in days.values() if v["supply"] and v["supply"] - v["sale"] < 0)
        rates = []
        for v in days.values():
            if v["supply"]:
                rates.append((v["supply"] - v["sale"]) / v["supply"] * 100.0)
        areas.append(
            {
                "area": area,
                "supply": round(supply, 2),
                "sale": round(sale, 2),
                "loss": round(loss, 2),
                "lossRate": None if rate is None else round(rate, 2),
                "days": len(days),
                "negativeDays": neg_days,
                "minDayRate": None if not rates else round(min(rates), 2),
                "maxDayRate": None if not rates else round(max(rates), 2),
            }
        )

    # ---- A1 一台区多终端
    n_term = params.get("n_term_multi")
    term_counts = []
    for area, info in sorted(detail.items()):
        term_counts.append(
            {
                "area": area,
                "terminals": sorted(info["terminals"]),
                "terminalCount": len(info["terminals"]),
                "totalMeterTerminals": sorted(info["totalTerminals"]),
                "userRows": info["users"],
            }
        )
    a1 = {
        "rule": "A1",
        "status": "GAP" if n_term in (None, "") else "OK",
        "threshold": n_term,
        "rows": term_counts,
        "hits": [],
        "note": "未传入 n_term_multi，只做终端数罗列，不判定异常"
        if n_term in (None, "")
        else f"按 n_term_multi={n_term} 判定",
    }
    if n_term not in (None, ""):
        a1["hits"] = [r["area"] for r in term_counts if r["terminalCount"] >= int(n_term)]
    else:
        gaps.append("A1 判定下限 n_term_multi 未提供：仅输出终端计数，不产生异常结论")

    # ---- A3 户表终端与总表终端不一致
    a3 = {"rule": "A3", "status": "OK", "rows": [], "note": ""}
    if not tg_flag:
        a3["status"] = "GAP"
        a3["note"] = "缺少总表标识列（Is Tg Meter），无法确定台区总表所在终端"
        gaps.append("A3 需要总表标识列（Is Tg Meter）才能核对终端归属")
    elif not roles.get("terminal"):
        a3["status"] = "GAP"
        a3["note"] = "缺少终端列"
        gaps.append("A3 需要终端地址码列")
    else:
        for row in rows:
            area = area_key(row)
            flag = norm(row.get(tg_flag))
            if flag in ("是", "Y", "y", "1", "True", "true"):
                continue
            terminal = norm(row.get(roles["terminal"]))
            info = detail.get(area, {})
            totals = info.get("totalTerminals") or set()
            if terminal and totals and terminal not in totals:
                a3["rows"].append({"area": area, "terminal": terminal, "totalMeterTerminals": sorted(totals)})
        a3["count"] = len(a3["rows"])
        a3["note"] = "户表采集终端不在本台区总表终端集合内"

    # ---- A2 一终端多台区（终端台账的跨台区数）
    n_tai = params.get("n_tai_multi")
    a2 = {"rule": "A2", "status": "OK", "threshold": n_tai, "rows": [], "hits": []}
    if roles.get("crossCount"):
        cc = roles["crossCount"]
        for row in rows:
            val = to_num(row.get(cc))
            if val is None:
                continue
            a2["rows"].append(
                {
                    "terminal": norm(row.get(roles.get("terminal"))) or "—",
                    "crossCount": val,
                    "areas": norm(row.get("覆盖台区")),
                }
            )
        if n_tai in (None, ""):
            a2["status"] = "GAP"
            a2["note"] = "未传入 n_tai_multi，只做跨台区数罗列，不判异常"
            gaps.append("A2 判定下限 n_tai_multi 未提供：仅输出跨台区数，不产生异常结论")
        else:
            a2["hits"] = [r["terminal"] for r in a2["rows"] if r["crossCount"] >= int(n_tai)]
            a2["note"] = f"按 n_tai_multi={n_tai} 判定"
    else:
        a2["status"] = "GAP"
        a2["note"] = "缺少跨台区数列，无法执行"
        gaps.append("A2 需要终端台账的跨台区数列")

    # ---- B2 负损（定义性，不需要阈值）
    neg = [a for a in areas if a["negativeDays"] > 0]
    b2 = {
        "rule": "B2",
        "status": "OK",
        "rows": neg,
        "note": "负损 = 该周期线损电量 < 0（售电量 > 供电量），无容差参数"
        if params.get("th_loss_neg") in (None, "")
        else f"应用容差 th_loss_neg={params.get('th_loss_neg')}",
    }
    if params.get("th_loss_neg") in (None, ""):
        gaps.append("B2 负损容差 th_loss_neg 未提供，按 0 容差列事实")

    # ---- B1 高损（需要阈值）
    th_high = params.get("th_loss_high")
    b1 = {"rule": "B1", "status": "GAP" if th_high in (None, "") else "OK", "threshold": th_high, "hits": []}
    if th_high in (None, ""):
        gaps.append("B1 高损阈值 th_loss_high 未提供：只输出线损率分布，不判高损")
    else:
        b1["hits"] = [a["area"] for a in areas if a["lossRate"] is not None and a["lossRate"] > float(th_high)]

    # ---- 去重影响
    dedupe = {"groups": len(dup_keys), "rowsDropped": dropped}
    if dup_keys and has_supply and has_sale:
        supply_all = sum(
            to_num(r.get(roles["supply"])) or 0.0 for r in data["rows"]
        )
        sale_raw = sum(to_num(r.get(roles["sale"])) or 0.0 for r in data["rows"])
        sale_clean = sum(by_area_sale.values())
        if supply_all:
            dedupe["before"] = {
                "sale": round(sale_raw, 2),
                "lossRate": round((supply_all - sale_raw) / supply_all * 100.0, 2),
            }
            dedupe["after"] = {
                "sale": round(sale_clean, 2),
                "lossRate": round((supply_all - sale_clean) / supply_all * 100.0, 2),
            }
            dedupe["deltaLossRate"] = round(dedupe["after"]["lossRate"] - dedupe["before"]["lossRate"], 2)

    result = {
        "ok": True,
        "file": os.path.abspath(path),
        "sheet": data["sheet"],
        "headerRow": data["headerRow"],
        "rowCount": len(data["rows"]),
        "keptRows": len(rows),
        "mapping": roles,
        "mappingSource": mapping_source,
        "energyColumns": energies,
        "dateRange": _date_range(data["rows"], roles.get("date")),
        "params": params,
        "dedupe": dedupe,
        "areas": areas,
        "a1": a1,
        "a2": a2,
        "a3": a3,
        "b1": b1,
        "b2": b2,
        "gaps": gaps,
        "disclaimer": "本结果为数据诊断，异常结论需现场核实后方可用于处理。",
    }
    result["markdown"] = render_diagnose(result)
    return result


def render_diagnose(res):
    lines = [f"# 台区线损诊断 · {os.path.basename(res['file'])}", ""]
    lines.append(f"- 工作表：`{res['sheet']}`，表头第 {res['headerRow']} 行")
    lines.append(f"- 行数：{res['rowCount']} → 去重后 {res['keptRows']}")
    if res["dateRange"]:
        d = res["dateRange"]
        lines.append(f"- 周期：{d['min']} ~ {d['max']}（{d['distinct']} 个）")
    lines.append(f"- 字段映射：{res['mappingSource']}")
    lines.append("")
    ded = res["dedupe"]
    lines.append("## 表一 · 重复行")
    lines.append("")
    lines.append(f"- 重复组 {ded['groups']} 组，精简 {ded['rowsDropped']} 行")
    if "before" in ded:
        b, a = ded["before"], ded["after"]
        lines.append(f"- 未去重：售电 {b['sale']}，线损率 {b['lossRate']}%")
        lines.append(f"- 去重后：售电 {a['sale']}，线损率 {a['lossRate']}%（差 {ded['deltaLossRate']} 个百分点）")
    lines.append("")
    if res["areas"]:
        lines.append("## 表二 · 台区线损（重算）")
        lines.append("")
        lines.append("| 台区 | 供电量 | 售电量 | 线损电量 | 线损率% | 周期数 | 负损天数 | 日线损率区间% |")
        lines.append("|---|---|---|---|---|---|---|---|")
        for a in res["areas"]:
            rng = "—" if a["minDayRate"] is None else f"{a['minDayRate']} ~ {a['maxDayRate']}"
            lines.append(
                f"| {a['area']} | {a['supply']} | {a['sale']} | {a['loss']} | "
                f"{'—' if a['lossRate'] is None else a['lossRate']} | {a['days']} | {a['negativeDays']} | {rng} |"
            )
        lines.append("")
    lines.append("## 表三 · A1 一台区多终端")
    lines.append("")
    lines.append(f"- 状态：{res['a1']['status']}；{res['a1']['note']}")
    lines.append("")
    lines.append("| 台区 | 终端数 | 终端清单 | 总表终端 | 户表行数 |")
    lines.append("|---|---|---|---|---|")
    for r in res["a1"]["rows"]:
        lines.append(
            f"| {r['area']} | {r['terminalCount']} | {', '.join(r['terminals']) or '—'} | "
            f"{', '.join(r['totalMeterTerminals']) or '—'} | {r['userRows']} |"
        )
    if res["a1"]["hits"]:
        lines.append("")
        lines.append(f"- 命中（终端数 ≥ {res['a1']['threshold']}）：{', '.join(res['a1']['hits'])}")
    lines.append("")
    lines.append("## 表三b · A2 一终端多台区")
    lines.append("")
    lines.append(f"- 状态：{res['a2']['status']}；{res['a2'].get('note', '')}")
    if res["a2"]["rows"]:
        lines.append("")
        lines.append("| 终端 | 跨台区数 | 覆盖台区 |")
        lines.append("|---|---|---|")
        for r in res["a2"]["rows"][:15]:
            lines.append(f"| {r['terminal']} | {int(r['crossCount'])} | {r['areas']} |")
        if res["a2"]["hits"]:
            lines.append("")
            lines.append(f"- 命中（跨台区数 ≥ {res['a2']['threshold']}）：{', '.join(res['a2']['hits'])}")
    lines.append("")
    lines.append("## 表四 · A3 户表终端与总表终端不一致")
    lines.append("")
    lines.append(f"- 状态：{res['a3']['status']}；{res['a3']['note']}")
    if res["a3"].get("count"):
        lines.append(f"- 不一致户表行：{res['a3']['count']}")
        for r in res["a3"]["rows"][:20]:
            lines.append(f"  - {r['area']}：户表终端 {r['terminal']}，总表终端 {', '.join(r['totalMeterTerminals'])}")
        if res["a3"]["count"] > 20:
            lines.append(f"  - …余 {res['a3']['count'] - 20} 行")
    lines.append("")
    lines.append("## 表五 · 线损异常（B 类）")
    lines.append("")
    lines.append(f"- B1 高损：{res['b1']['status']}" + (f"，阈值 {res['b1']['threshold']}%，命中 {', '.join(res['b1']['hits']) or '无'}" if res["b1"]["status"] == "OK" else ""))
    neg_areas = ", ".join(a["area"] for a in res["b2"]["rows"]) or "无"
    lines.append(f"- B2 负损：{res['b2']['status']}，出现负损的台区：{neg_areas}")
    lines.append("")
    lines.append("## 表六 · 缺口")
    lines.append("")
    if res["gaps"]:
        for g in res["gaps"]:
            lines.append(f"- {g}")
    else:
        lines.append("- 无")
    lines.append("")
    lines.append(f"> {res['disclaimer']}")
    return "\n".join(lines)


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"ok": False, "error": "用法: engine.py <scan|diagnose> <json>"}, ensure_ascii=False))
        return
    mode = sys.argv[1]
    args = json.loads(sys.argv[2])
    try:
        result = scan(args) if mode == "scan" else diagnose(args)
    except Exception as exc:  # 把失败原因原样交给调用方
        result = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
