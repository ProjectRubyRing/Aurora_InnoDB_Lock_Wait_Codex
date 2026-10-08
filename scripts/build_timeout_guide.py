"""Build an explicitly formatted, source-linked XLSX without Excel automation."""

from __future__ import annotations

import json
import math
import re
from pathlib import Path

import xlsxwriter
from PIL import ImageFont

from guide_content import AS_OF, EFFECTIVE_CASES, ENVIRONMENT, INPUTS, SECTIONS, SOURCES


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"
BOOK_PATH = OUT / "JBossEAP81_Aurora84_Timeout_DeepDive.xlsx"
FONT_NAME = "Meiryo UI"
FONT_FILE = Path("C:/Windows/Fonts/meiryo.ttc")
FONT_BOLD_FILE = Path("C:/Windows/Fonts/meiryob.ttc")
WIDTHS = [8, 25, 38, 57, 44, 13]
COLOR = {
    "navy": "18344B", "teal": "006E78", "text": "223C4D",
    "muted": "536B7A", "line": "D9E3EA", "blue": "E8F1FC",
    "pale": "F3F7FA", "green": "ECF6EF", "amber": "FFF4DA",
    "red": "FCE8E6", "white": "FFFFFF", "input": "DCEBFF",
    "link": "145FA6", "purple": "675485", "commit": "2D805B",
}


def pixel_width(width: float) -> int:
    # Excel's column widths are based on the default digit width, not a Japanese
    # character count. Use the conventional 7-pixel digit width plus padding.
    return math.floor(width * 7 + 5)


def measured_font(size: float = 11, bold: bool = False):
    return ImageFont.truetype(str(FONT_BOLD_FILE if bold else FONT_FILE),
                              round(size * 96 / 72), index=2)


def wrap_lines(text: str, pixels: int, size: float = 11, bold: bool = False):
    """Conservative measured wrapping, with word-aware Latin token boundaries.

    Keep CLI/code/URL contents intact in the workbook. This routine predicts
    viewer line wrapping for row heights and the separate layout preview.
    Excel may wrap further; leave width and vertical padding for that case.
    """
    font = measured_font(size, bold)
    usable = max(18, pixels - 20)
    output = []
    for para in str(text).split("\n"):
        if not para:
            output.append("")
            continue
        tokens = re.findall(r"[A-Za-z0-9_./:=<>$@()+,;\-]+|\s+|.", para)
        line = ""
        for token in tokens:
            if font.getlength(line + token) <= usable:
                line += token
                continue
            if line:
                output.append(line.rstrip())
                line = ""
            # Long CLI names and URLs must still fit; split them by measured
            # characters instead of relying on the spreadsheet viewer.
            for char in token:
                if line and font.getlength(line + char) > usable:
                    output.append(line.rstrip())
                    line = ""
                line += char
        output.append(line.rstrip())
    return output or [""]


def formatted_text(text, width, size=11, bold=False):
    # Automatic wrapping is a cell style. Artificial line breaks inside a CLI
    # command would make copying the command from Excel unsafe or incorrect.
    return str(text)


def height_for(values, widths, size=11, bold_cols=()):
    lines = max(len(wrap_lines(str(v), pixel_width(w), size, i in bold_cols))
                for i, (v, w) in enumerate(zip(values, widths)))
    return max(38, lines * 17.5 + 14)


def effective_limit(jta: str, switch: str, remaining: float, q: int, p: int):
    if jta == "true" and switch == "true":
        if remaining <= 0:
            return "実行拒否"
        rounded = math.ceil(remaining)
        return min(rounded, p) if p > 0 else rounded
    return p if p > 0 else q


def budget_results(inputs):
    t, e, q, p, switch, jta, lock, socket, connect, pool, other, count, commit, cleanup, safety, outer, xa = inputs
    remaining = max(0, t - e)
    rounded = math.ceil(remaining)
    eff = effective_limit(jta, switch, remaining, q, p)
    reserve = max(commit, cleanup) + safety
    usable = max(0, math.floor(remaining - reserve))
    proposed = ("新SQLを開始しない" if usable < 1 else
                min(p, usable) if p > 0 else
                min(q, usable) if q > 0 else "個別SQL上限を設定")
    q_budget = p if p > 0 else (0 if switch == "true" and jta == "true" else q)
    full = pool / 1000 + connect / 1000 + other + count * q_budget + reserve if q_budget > 0 else "SQL上限未確定"
    head_t = t - full if isinstance(full, (float, int)) else "未判定"
    head_h = outer - full if isinstance(full, (float, int)) else "未判定"
    fixed_note = ("実行拒否" if eff == "実行拒否" else
                  "JTA残りが優先（Qは上限でない）" if switch == "true" and jta == "true" and p == 0 else
                  "正の個別Pを使用" if p > 0 else "固定Qを使用")
    lock_note = ("実行拒否" if not isinstance(eff, (int, float)) else
                 "SQL期限なし" if eff == 0 else
                 "行ロック待ちが短い" if lock < eff else "SQL実効値≦行ロック待ち")
    socket_note = ("実行拒否" if not isinstance(eff, (int, float)) else
                   "SQL期限なし" if eff == 0 else "socket期限なし" if socket == 0 else
                   "SQL予定値が短い" if eff < socket / 1000 else "socket設定を確認")
    result_note = ("SQL上限を確定して再計算" if not isinstance(head_t, (int, float)) else
                   "予算の余裕あり（実機未検証）" if head_t > 0 and head_h > 0 else "予算不足または余裕0")
    return [remaining, rounded, eff, fixed_note, reserve, usable, proposed, full,
            head_t, head_h, "非対応：RM期限を設定しない", lock_note, socket_note,
            "外側HはJTA Tより長い" if outer > t else "外側が先に終了する可能性", result_note]


class Builder:
    def __init__(self):
        OUT.mkdir(exist_ok=True)
        self.wb = xlsxwriter.Workbook(BOOK_PATH, {"strings_to_formulas": False, "strings_to_urls": False})
        self.wb.formats[0].set_font_name(FONT_NAME)
        self.wb.formats[0].set_font_size(11)
        self.wb.set_properties({"title": "JBoss EAP 8.1とAurora MySQLのタイムアウト徹底解説",
                                "subject": "JTA / XA / query-timeout / set-tx-query-timeout / Narayana 2PC",
                                "author": "Codex", "comments": f"Primary-source research as of {AS_OF}. Runtime tests not performed."})
        self.wb.set_calc_mode("auto")
        self.formats = {}
        self.layout = {}
        self.formula_records = []
        self.source_rows = {s[0]: i + 5 for i, s in enumerate(SOURCES)}

    def fmt(self, kind="body", stripe=0, **extra):
        key = (kind, stripe, tuple(sorted(extra.items())))
        if key in self.formats:
            return self.formats[key]
        props = {"font_name": FONT_NAME, "font_size": 11, "font_color": COLOR["text"],
                 "text_wrap": True, "valign": "top", "border": 1, "border_color": COLOR["line"],
                 "bg_color": COLOR["white"] if stripe == 0 else COLOR["pale"]}
        if kind == "title":
            props.update(font_size=21, bold=True, font_color=COLOR["white"], bg_color=COLOR["navy"], border=0, valign="vcenter")
        elif kind == "subtitle":
            props.update(font_size=11, font_color=COLOR["muted"], bg_color=COLOR["white"], border=0, valign="vcenter")
        elif kind == "banner":
            props.update(bg_color=COLOR["amber"], font_color=COLOR["navy"], border=0, valign="vcenter")
        elif kind == "header":
            props.update(bold=True, font_color=COLOR["white"], bg_color=COLOR["teal"], valign="vcenter")
        elif kind == "topic":
            props.update(bold=True, bg_color=COLOR["blue"] if stripe == 0 else "DEEBF8")
        elif kind == "easy":
            props.update(bg_color=COLOR["green"] if stripe == 0 else "E5F1E9")
        elif kind == "decision":
            props.update(bg_color=COLOR["amber"] if stripe == 0 else "F9EED4")
        elif kind == "source":
            props.update(font_size=10, font_color=COLOR["link"], underline=True)
        elif kind == "input":
            props.update(bg_color=COLOR["input"], font_color=COLOR["link"], bold=True, align="center", valign="vcenter", num_format="0.###")
        elif kind == "number":
            props.update(align="center", valign="vcenter", num_format="0.###")
        elif kind == "calc":
            props.update(bg_color=COLOR["green"], bold=True, num_format="0.###", valign="vcenter")
        elif kind == "nav":
            props.update(font_color=COLOR["link"], underline=True, border=0, bg_color=COLOR["white"], valign="vcenter")
        elif kind == "block":
            props.update(bg_color=COLOR["teal"], font_color=COLOR["white"], bold=True, align="center", valign="vcenter")
        props.update(extra)
        obj = self.wb.add_format(props)
        self.formats[key] = obj
        return obj

    def base(self, name, title, subtitle, widths=WIDTHS, zoom=85):
        ws = self.wb.add_worksheet(name)
        self.layout[name] = {"widths": widths, "rows": {}, "merged": []}
        ws.hide_gridlines(2)
        ws.set_zoom(zoom)
        ws.set_default_row(22)
        ws.set_tab_color(COLOR["teal"])
        for col, width in enumerate(widths):
            ws.set_column(col, col, width, self.fmt())
        ws.merge_range(0, 0, 0, len(widths) - 1, title, self.fmt("title"))
        ws.set_row(0, 45)
        self.record_merge(name, 0, 0, len(widths) - 1, title, 45, "title")
        ws.merge_range(1, 0, 1, len(widths) - 1, subtitle, self.fmt("subtitle"))
        ws.set_row(1, 37)
        self.record_merge(name, 1, 0, len(widths) - 1, subtitle, 37, "subtitle")
        ws.write_url(2, 0, "internal:'00_はじめに'!A1", self.fmt("nav"), "目次")
        metadata = f"対象：{ENVIRONMENT}  ／  調査基準日：{AS_OF}"
        metadata_width = sum(pixel_width(w) for w in widths[1:])
        metadata_height = max(29, len(wrap_lines(metadata, metadata_width)) * 17.5 + 10)
        ws.merge_range(2, 1, 2, len(widths) - 1, metadata, self.fmt("subtitle"))
        ws.set_row(2, metadata_height)
        self.record_merge(name, 2, 1, len(widths) - 1, metadata, metadata_height, "subtitle")
        ws.freeze_panes(5, 2)
        ws.set_landscape()
        ws.set_paper(8)  # A3; retain the full readable table width.
        ws.fit_to_pages(1, 0)
        ws.set_margins(0.28, 0.28, 0.38, 0.38)
        ws.repeat_rows(0, 4)
        ws.set_header('&L&"Meiryo UI,Regular"&9 JTA・XA・SQL timeout  |  EAP 8.1 / Aurora MySQL')
        ws.set_footer('&L&"Meiryo UI,Regular"&9 出典は21シート  |  実機試験は未実施&C&"Meiryo UI,Regular"&9&P / &N&R&"Meiryo UI,Regular"&9 ' + AS_OF)
        return ws

    def record_merge(self, name, row, first, last, text, height, kind):
        self.layout[name]["merged"].append({"row": row, "first": first, "last": last,
                                                 "text": text, "height": height, "kind": kind})

    def row(self, ws, index, values, kinds=None, widths=None, size=11, height=None):
        widths = widths or self.layout[ws.name]["widths"]
        kinds = kinds or ["body"] * len(values)
        bolds = [i for i, k in enumerate(kinds) if k in {"topic", "header", "calc", "input", "block"}]
        height = height or height_for(values, widths, size, bolds)
        if height > 390:
            raise ValueError(f"Split overlong row: {ws.name}, Excel row {index + 1}, {height}")
        ws.set_row(index, height)
        stored = []
        for col, (value, width, kind) in enumerate(zip(values, widths, kinds)):
            if isinstance(value, (int, float)):
                ws.write_number(index, col, value, self.fmt(kind, index % 2))
                rendered = str(value)
            else:
                col_size = 10 if kind == "source" else size
                rendered = formatted_text(value, width, col_size, col in bolds)
                ids = re.findall(r"S\d{2}", str(value)) if kind == "source" else []
                if ids:
                    dest = f"internal:'21_出典と確認範囲'!A{self.source_rows[ids[0]] + 1}"
                    ws.write_url(index, col, dest, self.fmt(kind, index % 2), rendered, "先頭の出典へ移動。全IDは出典一覧で確認。")
                else:
                    ws.write_string(index, col, rendered, self.fmt(kind, index % 2))
            stored.append({"column": col, "text": rendered, "kind": kind, "size": 10 if kind == "source" else size})
        self.layout[ws.name]["rows"][str(index)] = {"height": height, "cells": stored}

    def narrative(self, section):
        name, title, subtitle, rows = section
        ws = self.base(name, title, subtitle)
        headers = ["No.", "項目", "やさしい説明・たとえ", "しくみ・技術上の事実", "設定判断・具体例・確認事項", "根拠ID"]
        if name == "16_設定確認とCLI":
            headers[3] = "確認コマンド・設定例"
        elif name == "17_Java実装例":
            headers[3] = "コードの抜粋・しくみ"
        elif name == "19_検証ケース":
            headers[3] = "試験する条件・観察"
            headers[4] = "期待する観察・合格条件／未実施"
        self.row(ws, 4, headers, ["header"] * 6, height=40)
        for i, row in enumerate(rows, 5):
            self.row(ws, i, [i - 4, *row], ["number", "topic", "easy", "body", "decision", "source"])
        ws.autofilter(4, 0, len(rows) + 4, 5)
        ws.print_area(0, 0, len(rows) + 4, 5)
        if name in {"01_結論と設定案", "10_2PCと期限の境界", "14_設定値の候補"}:
            ws.set_tab_color(COLOR["navy"])

    def source_sheet(self):
        widths = [8, 21, 39, 53, 45, 46]
        ws = self.base("21_出典と確認範囲", "出典・確認日・適用範囲", "42件の一次資料。設定値の設計例・代表実装・実機未確認を明確に区別する。", widths, 80)
        ws.set_tab_color(COLOR["purple"])
        self.row(ws, 4, ["ID", "公開主体", "資料・固定版", "公式リンク（クリック）", "裏付ける内容", "適用範囲と限界"], ["header"] * 6, height=40)
        for i, source in enumerate(SOURCES, 5):
            sid, owner, title, url, evidence, scope = source
            values = [sid, owner, title, url, evidence, f"確認日：{AS_OF}\n{scope}"]
            self.row(ws, i, values, ["topic", "body", "body", "source", "body", "decision"], widths)
            display = formatted_text(url, widths[3], 10)
            ws.write_url(i, 3, url, self.fmt("source", i % 2), display, title)
        ws.autofilter(4, 0, len(SOURCES) + 4, 5)
        ws.print_area(0, 0, len(SOURCES) + 4, 5)

    def truth_sheet(self):
        widths = [8, 10, 12, 15, 15, 16, 16, 68]
        ws = self.base("09_実効値早見", "SQL timeoutの実効値を比べる", "新しいStatementの教育モデル。正の個別Pと固定Qを分ける。コード根拠：S05・S06・S07。", widths, 90)
        headers = ["例", "JTAあり", "set-tx", "残りR\n秒", "固定Q\n秒", "個別P\n秒", "実効値\n秒／拒否", "なぜこの値になるか"]
        self.row(ws, 4, headers, ["header"] * 8, height=48)
        for index, case in enumerate(EFFECTIVE_CASES, 5):
            self.row(ws, index, case, ["topic", "number", "number", "number", "number", "number", "calc", "body"], widths)
        notes = [
            "P=0は『新しいStatementに正の明示値なし』。Statement.setQueryTimeout(0)でEAPの自動設定を必ず解除できるという意味ではない。",
            "Rは上流TMの残り推定。端数切り上げは代表WrapperDataSourceの分岐を単純化したモデル。実TMの粒度・dispatch・取消遅れで実停止時刻は変わる。",
            "残り0の実行拒否はtrue＋JTAありの代表分岐。transactionの状態チェックが先に別の例外を返す場合もある。",
            "同じStatementの再利用では前回のtimeoutSetが次回に影響し得る。アプリが毎回正の個別Pを設定する方式と、close/cacheのresetを実機で検証する。",
            "q_effはdriverに設定する予定秒数。SQL完了、取消完了、rollback完了、lock解放の時刻を保証する値ではない。",
        ]
        start = len(EFFECTIVE_CASES) + 6
        for i, note in enumerate(notes, start):
            height = height_for([note], [sum(widths)])
            ws.merge_range(i, 0, i, 7, formatted_text(note, sum(widths)), self.fmt("banner"))
            ws.set_row(i, height)
            self.record_merge(ws.name, i, 0, 7, note, height, "banner")
        ws.autofilter(4, 0, len(EFFECTIVE_CASES) + 4, 7)
        ws.print_area(0, 0, start + len(notes) - 1, 7)

    def formula(self, ws, row, col, formula, value):
        ws.write_formula(row, col, formula, self.fmt("calc"), value)
        self.formula_records.append({"sheet": ws.name, "cell": xlsxwriter.utility.xl_rowcol_to_cell(row, col),
                                     "formula": formula, "cached_value": value})
        for entry in self.layout[ws.name]["rows"][str(row)]["cells"]:
            if entry["column"] == col:
                entry["text"] = str(value)

    def calculator(self):
        widths = [38, 18, 10, 36, 48, 13]
        ws = self.base("15_設定診断器", "設定値を変えて整合性を確認する", "青いB列＝入力、緑のE列＝計算。教育用モデル。実システムへの接続・設定変更は行わない。", widths, 90)
        ws.set_tab_color(COLOR["link"])
        self.row(ws, 4, ["入力項目", "変更する値", "単位", "計算・診断", "結果（自動計算）", "単位"], ["header"] * 6, height=40)
        values = [row[1] for row in INPUTS]
        cached = budget_results(values)
        output_specs = [
            ("JTA残りR（モデル）", '=MAX(0,$B$6-$B$7)', "秒"),
            ("Rの整数秒への切り上げ", '=ROUNDUP(E6,0)', "秒"),
            ("SQL実効timeout q_eff", '=IF(AND($B$10="true",$B$11="true"),IF(E6<=0,"実行拒否",IF($B$9>0,MIN(E7,$B$9),E7)),IF($B$9>0,$B$9,$B$8))', "秒/拒否"),
            ("固定Qと個別Pの扱い", '=IF(E8="実行拒否","実行拒否",IF(AND($B$10="true",$B$11="true",$B$9=0),"JTA残りが優先（Qは上限でない）",IF($B$9>0,"正の個別Pを使用","固定Qを使用")))', "説明"),
            ("終了用reserve K", '=MAX($B$18,$B$19)+$B$20', "秒"),
            ("reserve後のSQL可用秒数", '=MAX(0,ROUNDDOWN(E6-E10,0))', "秒"),
            ("業務側の次SQL予算", '=IF(E11<1,"新SQLを開始しない",IF($B$9>0,MIN($B$9,E11),IF($B$8>0,MIN($B$8,E11),"個別SQL上限を設定")))', "秒/判断"),
            ("業務全体の予算合計", '=IF(IF($B$9>0,$B$9,IF(AND($B$10="true",$B$11="true"),0,$B$8))>0,$B$15/1000+$B$14/1000+$B$16+$B$17*IF($B$9>0,$B$9,$B$8)+E10,"SQL上限未確定")', "秒"),
            ("JTA予算との差 T−合計", '=IF(ISNUMBER(E13),$B$6-E13,"未判定")', "秒"),
            ("外側予算との差 H−合計", '=IF(ISNUMBER(E13),$B$21-E13,"未判定")', "秒"),
            ("Connector/J9.7.0のXA値", None, "常に非対応"),
            ("行ロック待ちとの比較", '=IF(ISNUMBER(E8),IF(E8=0,"SQL期限なし",IF($B$12<E8,"行ロック待ちが短い","SQL実効値≦行ロック待ち")),"実行拒否")', "予定値"),
            ("socket設定との比較", '=IF(ISNUMBER(E8),IF(E8=0,"SQL期限なし",IF($B$13=0,"socket期限なし",IF(E8<$B$13/1000,"SQL予定値が短い","socket設定を確認"))),"実行拒否")', "予定値"),
            ("外側HとJTA Tの比較", '=IF($B$21>$B$6,"外側HはJTA Tより長い","外側が先に終了する可能性")', "説明"),
            ("予算モデルの総合判定", '=IF(ISNUMBER(E14),IF(AND(E14>0,E15>0),"予算の余裕あり（実機未検証）","予算不足または余裕0"),"SQL上限を確定して再計算")', "実機未検証"),
        ]
        for i, entry in enumerate(INPUTS, 5):
            label, value, unit, _ = entry
            out = output_specs[i - 5] if i - 5 < len(output_specs) else ("", None, "")
            out_value = cached[i - 5] if i - 5 < len(cached) else ""
            self.row(ws, i, [label, value, unit, out[0], out_value, out[2]],
                     ["topic", "input", "number", "topic", "calc", "number"], widths, height=66)
            if out[1]:
                self.formula(ws, i, 4, out[1], out_value)
            cell = f"B{i + 1}"
            if isinstance(value, str):
                options = {"validate": "list", "source": ["true", "false"]}
            else:
                options = {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 86400000 if unit == "ms" else 86400}
                if label in {"JTAの業務予算 T", "行ロック待ち L", "SQLの本数 N"}:
                    options["minimum"] = 1
                if label == "経過時間 e":
                    options["validate"] = "decimal"
                if label == "SQLの本数 N":
                    options["maximum"] = 10000
            options.update(input_title="教育用の設定入力", input_message=entry[3][:220],
                           error_title="入力の型・範囲を確認", error_message="指定された型・範囲内の値を入力してください。", error_type="stop")
            ws.data_validation(cell, options)
        ws.conditional_format("E14:E15", {"type": "cell", "criteria": "<=", "value": 0, "format": self.fmt(bg_color=COLOR["red"], bold=True)})
        ws.conditional_format("E9", {"type": "text", "criteria": "containing", "value": "Qは上限でない", "format": self.fmt(bg_color=COLOR["amber"], bold=True)})
        ws.conditional_format("E20", {"type": "text", "criteria": "containing", "value": "予算不足", "format": self.fmt(bg_color=COLOR["red"], bold=True)})
        notes = [
            ("初期値の読み方", "R52、q_eff10、reserve20、全体予算52、JTAとの差8秒。すべて説明用の候補で、現在の実環境値ではない。"),
            ("落とし穴の操作例", "個別P（B9）を0にすると、true＋JTAありではq_eff52になり、固定Q10は上限にならない。全体予算はSQL上限未確定になる。"),
            ("不足の操作例", "SQL本数N（B17）を3にすると、全体予算62秒、JTAとの差−2秒。Q/PがTより小さくても全体予算が不足する。"),
            ("期限直前の操作例", "経過e（B7）を55にするとR5、q_eff5、reserve後0。次SQLは開始しない。経過を60にするとtrue＋JTAありのSQL実行を拒否するモデル。"),
            ("JTAなしの扱い", "参加中=falseならq_effはPまたはQ。Rと業務reserveの欄はアプリ予算モデルとして残るが、実JTAが存在するという意味ではない。"),
            ("予算式の前提", "N本直列、物理接続作成1回、pool待ち1回、SQL外A、max(2PC,取消＋rollback)+余裕の保守モデル。実経路の重複は調整する。"),
            ("保証できないこと", "socketは総経過期限ではなく、SQL取消・rollback・2PC/recoveryの完了時刻も保証しない。『予算余裕あり』は実機試験合格を意味しない。"),
            ("正の個別SQL予算", "次SQL予算はreserveを残すための業務側の設定候補。これを自動でStatementへ書く実装は行わない。Pなしなら、候補を正のPとして明示し検証する。"),
            ("モデル外の状態", "TMから残り取得不可、既存StatementのtimeoutSet、rollback-only、CMT境界、ORM変換、異なるRMは計算器で再現しない。08/19で確認する。"),
        ]
        self.row(ws, 24, ["", "確認のポイント", "", "説明", "", ""], ["header"] * 6, height=38)
        for i, (label, note) in enumerate(notes, 25):
            self.row(ws, i, [i - 24, label, "", note, "計算式は編集不要。青い入力セルのみ変更。", "S05 S06"],
                     ["number", "topic", "body", "body", "decision", "source"], widths)
        ws.print_area(0, 0, 24 + len(notes), 5)

    def home(self):
        ws = self.base("00_はじめに", "JTA・XA・SQLタイムアウト 徹底解説", "小学生向けのたとえから、Narayanaの2PC・Connector/J実装・設定設計まで。", WIDTHS, 90)
        ws.set_tab_color(COLOR["navy"])
        self.row(ws, 4, ["No.", "最初に読むこと", "やさしい説明", "この資料の前提・確認範囲", "使い方", "根拠"], ["header"] * 6, height=40)
        intro = [
            ("最重要の答え", "2PCを使っても、全体の時計JTAは必要です。", "JTA期限は業務寿命に必要。しかしXA prepare/commitとrecoveryの厳密な完了上限にはならない。", "まず01、次に10と22の図解を読む。", "S18 S19 S25"),
            ("今回の環境", "指定された道具の型番を対象にします。", ENVIRONMENT + "。利用者回答による指定。現在の4設定値とEAP update番号は未確認。", "9.7.0の固定タグとEAP公開代表ビルドのコードを照合。", "S04 S10 S31 S32"),
            ("名称の訂正", "resource＝担当係、resume＝再開。", "依頼のxa-resume-timeoutは、利用者回答でxa-resource-timeoutを指すと確認。正式名を以後の説明とCLIに使用。", "Auroraの自動再開やJTA suspend/resumeと区別。", "S01 S28 S33"),
            ("検証候補の主案", "全体60秒、一問個別10秒、片付けも残します。", "JTA60秒、XA0、固定Q10秒、set-tx=true＋正の個別P10秒。候補値は製品公式推奨の唯一解ではない。", "14で条件を確認し、15の青いセルで整合性を計算。", "S05 S10 / 設計例"),
            ("実施した作業の範囲", "説明書とソースを調べ、Excelの文字と計算を検証します。", "公式資料・固定タグの調査、Excel構造・書式・リンク・数式と保存済み結果の検証。実EAP/DB設定変更や動作試験は未実施。", "19は未実施の試験計画。実機での合格済み一覧ではない。", "作業範囲"),
            ("読みやすさの設定", "見出しを固定し、長い文は折り返して表示します。", "全セルMeiryo UI。本文11pt。青緑の見出し、緑のたとえ、黄の判断、根拠リンク。列幅・行高は実フォントの文字幅から設定。", "シート目次、フィルター、ウィンドウ枠固定。印刷はA3横・1ページ幅・縦は複数ページ。", "書式設定"),
        ]
        for i, data in enumerate(intro, 5):
            self.row(ws, i, [i - 4, *data], ["number", "topic", "easy", "body", "decision", "source"])
        start = 13
        self.row(ws, start, ["No.", "クリックして移動", "読みたい内容", "説明・位置づけ", "おすすめ順", ""], ["header"] * 6, height=38)
        sheet_info = {s[0]: (s[1], s[2]) for s in SECTIONS}
        sheet_info.update({"09_実効値早見": ("実効timeoutの14例", "QとP、JTA残りの優先順位を具体的な数字で比較。"),
                           "15_設定診断器": ("入力を変えて計算", "青いセルを編集してq_eff、reserve、業務全体予算を確認。"),
                           "21_出典と確認範囲": ("42件の一次資料", "公式リンク、固定版、確認日、適用限界。"),
                           "22_動作イメージ": ("時計と2PCの図解", "正常／期限直前／prepare／結果不明を色と矢印で理解。")})
        for row, name in enumerate(sorted(sheet_info), start + 1):
            title, desc = sheet_info[name]
            step = "① 基礎" if int(name[:2]) <= 4 else "② 原理" if int(name[:2]) <= 12 else "③ 設定・運用"
            if name.startswith("22"):
                step = "① 最初に見る図解"
            self.row(ws, row, [row - start, name, title, desc, step, ""], ["number", "topic", "easy", "body", "decision", "body"])
            ws.write_url(row, 1, f"internal:'{name}'!A1", self.fmt("source", row % 2, font_size=11, bold=True), formatted_text(name, WIDTHS[1], bold=True))
        ws.print_area(0, 0, start + len(sheet_info), 5)

    def diagram(self):
        widths = [31] + [13] * 10
        ws = self.base("22_動作イメージ", "4つの時計と2PCの境界を図で見る", "矢印・色は説明用。時刻は縮尺を含めた教育例で、実測結果ではない。", widths, 90)
        ws.set_tab_color(COLOR["commit"])
        ws.freeze_panes(5, 1)
        self.row(ws, 4, ["A：通常の処理と予算", *[f"{t}秒" for t in range(0, 100, 10)]], ["header"] * 11, height=38)
        rows = [
            ("JTAの活動期限：T60", {i: ("全体時計", "blue") for i in range(1, 7)}),
            ("SQL1：例 5〜15秒", {1: ("開始 →", "green"), 2: ("終了", "green")}),
            ("SQL2：例 25〜35秒", {3: ("開始 →", "green"), 4: ("終了", "green")}),
            ("終了用reserve：20秒", {5: ("余裕を確保", "amber"), 6: ("余裕を確保", "amber")}),
            ("2PC：例 40〜45秒", {5: ("prepare\n決定→commit", "green")}),
            ("外側応答予算：H90", {i: ("request予算", "pale") for i in range(1, 10)}),
        ]
        for row, (label, blocks) in enumerate(rows, 5):
            self.row(ws, row, [label] + [""] * 10, ["topic"] + ["body"] * 10, widths, height=66)
            for col, (text, color) in blocks.items():
                ws.write_string(row, col, text, self.fmt("body", bg_color=COLOR[color], align="center", valign="vcenter", bold=True))
                self.layout[ws.name]["rows"][str(row)]["cells"][col]["text"] = text
        self.wide_note(ws, 12, "全体の残りへSQLを連動しても、終了用reserve20秒は自動では差し引かない。上の2PC位置は正常例の予定で、JTA60秒以内の通信完了を保証しない。", widths)
        self.row(ws, 14, ["B：2PCとJTAの境界", "begin", "SQL", "flush", "XA END", "prepare", "決定記録", "commit", "応答", "障害後", "recovery"], ["header"] * 11, height=48)
        state = ["TMの状態（単純化）", "RUNNING", "RUNNING", "RUNNING", "状態確認", "PREPARING", "決定境界", "COMMITTING", "終了", "未解決あり", "同じ結果へ"]
        self.row(ws, 15, state, ["topic"] + ["body"] * 10, widths, height=76)
        clocks = ["JTAの位置づけ", "寿命を監視", "寿命を監視", "競合し得る", "状態次第", "強制終了を\n保証しない", "逆の決定へ\n変えない", "活動期限で\n戻せない", "完了を確認", "期限経過≠\n解決済み", "別の時間軸"]
        self.row(ws, 16, clocks, ["topic"] + ["decision"] * 10, widths, height=86)
        sql = ["SQL query timeout", "対象外", "対象", "SQLなら対象", "内部XA\n自動対象外", "内部XA\n自動対象外", "対象外", "内部XA\n自動対象外", "fetchは確認", "通信を確認", "復旧接続\n別途確認"]
        self.row(ws, 17, sql, ["topic"] + ["easy"] * 10, widths, height=86)
        self.wide_note(ws, 19, "核心：代表Narayanaのrunning条件はRUNNING/ABORT_ONLY。Reaper登録がcommit完了まで残っていても、prepare/commitをJTA期限で強制終了できる意味ではない。commit決定後は同じ結果へ回復する。根拠：S19〜S23。", widths)
        self.row(ws, 21, ["C：固定Qと個別P", "JTA残り", "固定Q", "個別P", "set-tx", "実効SQL", "", "読むポイント", "", "", ""], ["header"] * 11, height=38)
        cases = [
            ["落とし穴", "50秒", "10秒", "なし", "true", "50秒", "", "Q10はminに", "入らない", "", ""],
            ["主案", "50秒", "10秒", "10秒", "true", "10秒", "", "個別Pと残り", "の短い方", "", ""],
            ["残り僅か", "4秒", "10秒", "10秒", "true", "4秒", "", "後始末時間は", "別に確保", "", ""],
        ]
        for row, values in enumerate(cases, 22):
            self.row(ws, row, values, ["topic", "number", "number", "number", "number", "calc", "body", "body", "body", "body", "body"], widths, height=66)
        self.wide_note(ws, 26, "D：commit応答消失  ｜  TM：prepare成功 → commit決定を保存 → XA COMMIT送信 → 応答待ちtimeout  ｜  RM：受信 → commit済み → 応答だけ消失もある。『timeout＝未反映』と判断せず、Xid・取引ID・recoveryで最終結果を確かめる。根拠：S10・S13・S25。", widths)
        self.wide_note(ws, 28, "色の意味：青＝全体予算、緑＝SQL/正常確定、黄＝判断・終了予備、淡色＝外側request。時刻セルは10秒区間の概略で、SQL1/2や2PCの正確な縮尺を保証するグラフではない。11の詳細時系列も参照。", widths)
        ws.print_area(0, 0, 28, 10)

    def wide_note(self, ws, row, text, widths):
        total = sum(widths)
        h = height_for([text], [total]) + 9
        ws.merge_range(row, 0, row, len(widths) - 1, formatted_text(text, total), self.fmt("banner"))
        ws.set_row(row, h)
        self.record_merge(ws.name, row, 0, len(widths) - 1, text, h, "banner")

    def run(self):
        builders = {"00_はじめに": self.home, "09_実効値早見": self.truth_sheet,
                    "15_設定診断器": self.calculator, "21_出典と確認範囲": self.source_sheet,
                    "22_動作イメージ": self.diagram}
        sections = {s[0]: s for s in SECTIONS}
        for name in sorted(set(builders) | set(sections)):
            if name in builders:
                builders[name]()
            else:
                self.narrative(sections[name])
        self.wb.worksheets()[0].activate()
        self.wb.worksheets()[0].set_first_sheet()
        self.wb.close()
        manifest = {"workbook": BOOK_PATH.name, "as_of": AS_OF,
                    "environment": ENVIRONMENT, "font": FONT_NAME,
                    "narrative_rows": sum(len(s[3]) for s in SECTIONS),
                    "sources": len(SOURCES), "formulas": self.formula_records, "layout": self.layout}
        (OUT / "workbook_validation_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"workbook": str(BOOK_PATH), "sheets": len(self.layout),
                          "narrative_rows": manifest["narrative_rows"], "sources": len(SOURCES),
                          "bytes": BOOK_PATH.stat().st_size}, ensure_ascii=True))


if __name__ == "__main__":
    Builder().run()
