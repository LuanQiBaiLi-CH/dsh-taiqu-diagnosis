#!/usr/bin/env python3
"""engine.py 的回归测试：合成工作簿，无任何真实客户数据。

跑法： python3 tests/test_engine.py
退出码 0 = 全部通过。
"""
import json
import os
import subprocess
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.join(os.path.dirname(HERE), "lib", "engine.py")

FAILS = []


def check(name, cond, detail=""):
    if cond:
        print(f"  ok   {name}")
    else:
        FAILS.append(f"{name} {detail}")
        print(f"  FAIL {name} {detail}")


def run(mode, args):
    proc = subprocess.run(
        [sys.executable, ENGINE, mode, json.dumps(args, ensure_ascii=False)],
        capture_output=True,
        text=True,
    )
    if not proc.stdout.strip():
        raise AssertionError(f"引擎无输出：{proc.stderr[:400]}")
    return json.loads(proc.stdout)


def build_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "台区线损日报"
    ws.append(["台区编号", "台区名称", "日期", "供电量", "售电量"])
    days = ["2026-10-01", "2026-10-02", "2026-10-03"]
    for day in days:
        ws.append(["A", "A区", day, 1000, 950])   # 5.0%
        ws.append(["B", "B区", day, 1000, 1050])  # -5.0%，三天全负损
        ws.append(["C", "C区", day, 800, 760])    # 5.0%

    ws2 = wb.create_sheet("用户日电量明细")
    ws2.append(["台区编号", "终端地址码", "电表条码", "测量点号", "数据日期", "客户编号"])
    ws2.append(["A", "T1", "E1", "1", "2026-10-01", "U1"])
    ws2.append(["A", "T1", "E1", "", "2026-10-01", "U1"])   # 重复行，测量点号为空
    ws2.append(["A", "T1", "E2", "2", "2026-10-01", "U2"])
    ws2.append(["A", "T2", "E3", "3", "2026-10-01", "U3"])
    ws2.append(["A", "T2", "E4", "4", "2026-10-02", "U4"])
    wb.save(path)


def main():
    tmpdir = tempfile.mkdtemp(prefix="taiqu-test-")
    book = os.path.join(tmpdir, "synthetic.xlsx")
    build_workbook(book)

    print("scan")
    scan = run("scan", {"file": book, "sheet": "台区线损日报"})
    check("scan 成功", scan.get("ok") is True)
    check("表头行 = 1", scan.get("headerRow") == 1, scan.get("headerRow"))
    check("角色映射含台区编号", scan["mapping"].get("tgId") == "台区编号", scan["mapping"])
    check("日期范围 3 个周期", (scan.get("dateRange") or {}).get("distinct") == 3)

    print("diagnose / 线损重算")
    diag = run(
        "diagnose",
        {"file": book, "sheet": "台区线损日报", "params": {"th_loss_high": 4.9}},
    )
    rates = {a["area"]: a["lossRate"] for a in diag["areas"]}
    check("A 区 5.0%（台区以编号为准）", rates.get("A") == 5.0, rates)
    check("B 区 -5.0%", rates.get("B") == -5.0, rates)
    check("C 区 5.0%", rates.get("C") == 5.0, rates)
    check("B 区 3 天全负损", next(a for a in diag["areas"] if a["area"] == "B")["negativeDays"] == 3)
    check("高损命中 A、C", diag["b1"]["hits"] == ["A", "C"], diag["b1"]["hits"])

    diag_hi = run(
        "diagnose",
        {"file": book, "sheet": "台区线损日报", "params": {"th_loss_high": 5.1}},
    )
    check("阈值 5.1 不命中", diag_hi["b1"]["hits"] == [], diag_hi["b1"]["hits"])

    diag_gap = run("diagnose", {"file": book, "sheet": "台区线损日报"})
    check("无阈值时 B1 = GAP", diag_gap["b1"]["status"] == "GAP")
    check("缺口清单非空", len(diag_gap["gaps"]) >= 1)

    print("diagnose / 去重与 A1")
    ded = run("diagnose", {"file": book, "sheet": "用户日电量明细"})
    check("重复组 = 1", ded["dedupe"]["groups"] == 1, ded["dedupe"])
    check("精简 1 行", ded["dedupe"]["rowsDropped"] == 1, ded["dedupe"])
    check("去重后行数 = 4", ded["keptRows"] == 4, ded["keptRows"])
    a1 = ded["a1"]
    check("A 区终端数 = 2", a1["rows"][0]["terminalCount"] == 2, a1["rows"])
    check("无下限时 A1 = GAP", a1["status"] == "GAP")
    check("A3 因缺总表标识报缺口", ded["a3"]["status"] == "GAP")

    ded_hit = run(
        "diagnose",
        {"file": book, "sheet": "用户日电量明细", "params": {"n_term_multi": 2}},
    )
    check("A1 命中 A 区", ded_hit["a1"]["hits"] == ["A"], ded_hit["a1"]["hits"])

    print("mapping 覆盖")
    over = run(
        "diagnose",
        {
            "file": book,
            "sheet": "台区线损日报",
            "mapping": {"supply": "供电量", "sale": "售电量"},
            "params": {"th_loss_high": 4.9},
        },
    )
    check("覆盖后仍命中 A、C", over["b1"]["hits"] == ["A", "C"], over["b1"]["hits"])

    print()
    if FAILS:
        print(f"{len(FAILS)} 项失败")
        for item in FAILS:
            print(" -", item)
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
