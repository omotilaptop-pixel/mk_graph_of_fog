# -*- coding: utf-8 -*-
"""
気象庁ダウンロード形式データ（只見_2026年1_5月_気象データ.xlsx など）を
自動で読み込んでグラフ化するプログラム。

【読み込む気象データ】
  B列: 気温(℃)
  C列: 降水量(mm)
  D列: 風速(m/s)
  G列: 相対湿度(％)
  H列: 露点温度(℃)

【J〜AP列（現象コード）】
  "/"      = その時刻に現象なし
  空白     = まだデータ未入力
  1〜10    = 現象発生コード（下記の意味）
      1  薄い川霧
      2  川霧
      3  濃い川霧
      4  薄い全体霧
      5  全体霧
      6  全体濃い霧
      7  薄い層雲
      8  濃い層雲
      9  霧雨
      10 雨

【出力グラフ（5項目 × 月ごとに分割）】
  ① 気温 × 現象コード（「/」・1〜10）
  ② 降水量 × 現象コード
  ③ 風速 × 現象コード
  ④ 相対湿度 × 現象コード
  ⑤ 露点温度 × 現象コード

  各グラフは、気象データの時系列（線 or 棒グラフ）と、
  その時刻に記録された現象コード（「/」・1〜10）を“レーン”形式で
  分けて表示する2段構成です。データ期間が長い場合に1枚の画像へ
  詰め込みすぎて見づらくならないよう、「年-月」ごとにファイルを
  分割して出力します（例: 5ヶ月分のデータなら、1項目につき5枚、
  合計25枚のPNGが生成されます）。

使い方:
    python3 weather_visualizer.py 入力ファイル.xlsx [出力フォルダ]

Jupyter / Google Colab で直接セルに貼って実行する場合は、
下の DEFAULT_INPUT_FILE / DEFAULT_OUTPUT_DIR を編集してください。
（ファイルが見つからない場合、Colabなら自動でアップロード画面が出ます）
"""

import sys
import os
import re
import glob
import subprocess

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.font_manager as fm
from matplotlib.patches import Patch
from matplotlib.colors import to_rgba
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


# ---------------------------------------------------------------------------
# 0. 日本語フォントの自動設定（文字化け対策）
# ---------------------------------------------------------------------------
_CJK_FONT_KEYWORDS = [
    "noto sans cjk", "noto serif cjk", "ipaex", "ipagothic", "ipa gothic",
    "takao", "vl gothic", "yu gothic", "ms gothic", "hiragino",
    "source han sans", "droid sans fallback",
]


def _find_cjk_font():
    """システムにインストール済みの日本語(CJK)対応フォント名を探す"""
    for f in fm.fontManager.ttflist:
        if any(k in f.name.lower() for k in _CJK_FONT_KEYWORDS):
            return f.name
    return None


def _install_noto_cjk_font():
    """(Colab/Linux環境向け) Noto Sans CJKフォントをaptで自動インストールする"""
    try:
        subprocess.run(["apt-get", "update"], stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL, timeout=120, check=False)
        result = subprocess.run(
            ["apt-get", "install", "-y", "fonts-noto-cjk"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300, check=False,
        )
        return result.returncode == 0
    except Exception:
        return False


def setup_japanese_font():
    """日本語フォントを自動検出し、見つからなければ自動インストールする"""
    font_name = _find_cjk_font()

    if font_name is None:
        print("日本語フォントが見つからないため、自動インストールを試みます…")
        if _install_noto_cjk_font():
            for fp in glob.glob("/usr/share/fonts/**/*.[ot]t[fc]", recursive=True):
                try:
                    fm.fontManager.addfont(fp)
                except Exception:
                    pass
            font_name = _find_cjk_font()

    if font_name:
        plt.rcParams["font.family"] = font_name
        print(f"日本語フォントを設定しました: {font_name}")
    else:
        print("【警告】日本語フォントの自動設定に失敗しました。文字が□になる場合があります。")
        print("Colabの場合、別セルで下記を実行後、ランタイムを再起動してから再実行してください:")
        print("  !apt-get -y install fonts-noto-cjk")

    plt.rcParams["axes.unicode_minus"] = False
    return font_name


setup_japanese_font()


# ---------------------------------------------------------------------------
# 1. 列レイアウト・現象コードの定義
# ---------------------------------------------------------------------------

COL_DATETIME = "A"
COL_TEMP = "B"        # 気温(℃)
COL_PRECIP = "C"      # 降水量(mm)
COL_WIND = "D"        # 風速(m/s)
COL_HUMID = "G"       # 相対湿度(％)
COL_DEWPOINT = "H"    # 露点温度(℃)
PHENOMENA_RANGE = ("J", "AP")  # 現象コード列の範囲

MAIN_COLUMNS = {
    COL_TEMP: "気温(℃)",
    COL_PRECIP: "降水量(mm)",
    COL_WIND: "風速(m/s)",
    COL_HUMID: "相対湿度(％)",
    COL_DEWPOINT: "露点温度(℃)",
}

# 現象コードの意味
PHENOM_LABELS = {
    1: "薄い川霧",
    2: "川霧",
    3: "濃い川霧",
    4: "薄い全体霧",
    5: "全体霧",
    6: "全体濃い霧",
    7: "薄い層雲",
    8: "濃い層雲",
    9: "霧雨",
    10: "雨",
}

# 現象コードごとの色（系統で色をまとめて区別しやすくする）
PHENOM_COLORS = {
    1: "#aed6f1",   # 薄い川霧   (薄い青)
    2: "#3498db",   # 川霧       (青)
    3: "#1a5276",   # 濃い川霧   (濃い青)
    4: "#dcdde1",   # 薄い全体霧 (薄い灰)
    5: "#909497",   # 全体霧     (灰)
    6: "#2c3e50",   # 全体濃い霧 (濃い灰/黒)
    7: "#f8c471",   # 薄い層雲   (薄い橙)
    8: "#d35400",   # 濃い層雲   (濃い橙)
    9: "#a9dfbf",   # 霧雨       (薄緑)
    10: "#196f3d",  # 雨         (濃緑)
}

SLASH_COLOR = "#f9e79f"   # 「/」(現象なし) の背景色


# ---------------------------------------------------------------------------
# 2. データ読み込み
# ---------------------------------------------------------------------------

def find_header_row(ws, search_col="A", keyword="年月日時", max_search_rows=15):
    """「年月日時」が入っているヘッダー行番号(1始まり)を探す"""
    col_idx = column_index_from_string(search_col)
    for r in range(1, max_search_rows + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v == keyword:
            return r
    raise ValueError(f"ヘッダー行（{search_col}列に「{keyword}」）が見つかりませんでした。")


def load_weather_data(filepath, sheet_name=None):
    """
    Excelファイルを読み込み、以下の2つのDataFrameを返す。
      main_df  : 日時 + 気温/降水量/風速/相対湿度/露点温度
      phenom_df: 日時 + J〜AP列の現象コード（生の値のまま）
    """
    wb = load_workbook(filepath, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    header_row = find_header_row(ws, COL_DATETIME)
    data_start_row = header_row + 1
    dt_col_idx = column_index_from_string(COL_DATETIME)

    start_idx = column_index_from_string(PHENOMENA_RANGE[0])
    end_idx = column_index_from_string(PHENOMENA_RANGE[1])
    phenom_cols = [get_column_letter(c) for c in range(start_idx, end_idx + 1)]

    main_rows = []
    phenom_rows = []
    for r in range(data_start_row, ws.max_row + 1):
        date_val = ws.cell(row=r, column=dt_col_idx).value
        if date_val is None:
            continue

        mrow = {"datetime": date_val}
        for col_letter, label in MAIN_COLUMNS.items():
            cidx = column_index_from_string(col_letter)
            mrow[label] = ws.cell(row=r, column=cidx).value
        main_rows.append(mrow)

        prow = {"datetime": date_val}
        for cidx, col_letter in zip(range(start_idx, end_idx + 1), phenom_cols):
            prow[col_letter] = ws.cell(row=r, column=cidx).value
        phenom_rows.append(prow)

    main_df = pd.DataFrame(main_rows)
    main_df["datetime"] = pd.to_datetime(main_df["datetime"])
    for label in MAIN_COLUMNS.values():
        main_df[label] = pd.to_numeric(main_df[label], errors="coerce")
    main_df = main_df.sort_values("datetime").reset_index(drop=True)

    phenom_df = pd.DataFrame(phenom_rows)
    phenom_df["datetime"] = pd.to_datetime(phenom_df["datetime"])
    phenom_df = phenom_df.sort_values("datetime").reset_index(drop=True)

    return main_df, phenom_df, phenom_cols


# ---------------------------------------------------------------------------
# 3. 現象コードのエンコードと、時刻ごとの代表値の計算
# ---------------------------------------------------------------------------

def encode_phenomena_cell(value):
    """
    セルの値を数値にエンコードする。
      None / NaN  -> np.nan (未入力＝データなし)
      "/"          -> 0      (現象なし)
      "1" "7" 等   -> その数値 (現象コード。複数並記 "1 4 8" は最大値を採用)
    """
    if value is None:
        return np.nan
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s == "/":
        return 0.0
    nums = re.findall(r"\d+", s)
    if nums:
        return float(max(int(n) for n in nums))
    return np.nan


def compute_phenomena_status(phenom_df, phenom_cols):
    """
    各時刻について、J〜AP列(33列)の中から代表値を1つ求める。
      NaN -> 全列が未入力（その時刻はまだ何もデータ入力されていない）
      0   -> 「/」のみが記録されている（現象なし、と記録済み）
      1〜10 -> 記録された現象コードのうち最大値
               （複数の現象が同時記録されている場合は、最も値の大きい
               現象コードを代表として表示する）
    """
    encoded = phenom_df[phenom_cols].map(encode_phenomena_cell)
    return encoded.max(axis=1, skipna=True)


# ---------------------------------------------------------------------------
# 3.5 グラフの横幅設定（データ期間に応じて自動で横に広げる）
# ---------------------------------------------------------------------------

# 1インチあたり何日分を表示するか。値を小さくするほど横に広がり、
# 現象コードの線同士の間隔が広くなって見やすくなる。
DAYS_PER_INCH = 0.55  # 月ごと分割にしたので、1インチあたりの日数をさらに小さくして余裕を持たせる
MIN_FIG_WIDTH = 16    # 横幅の最小値（インチ）
MAX_FIG_WIDTH = 60    # 横幅の最大値（インチ、1ヶ月分なのでこれで十分）
SAVE_DPI = 150        # 保存時の解像度


def compute_fig_width(times):
    """データの期間（日数）に応じて、横に十分広いfigsize幅を計算する"""
    total_days = (times.max() - times.min()).total_seconds() / 86400.0
    width = max(total_days, 1.0) / DAYS_PER_INCH
    return float(np.clip(width, MIN_FIG_WIDTH, MAX_FIG_WIDTH))


def split_by_month(df):
    """
    DataFrame（datetime列を含む）を「年-月」ごとに分割し、
    [(年月文字列, 部分DataFrame), ...] のリストを返す。
    """
    df = df.copy()
    df["__ym"] = df["datetime"].dt.strftime("%Y-%m")
    groups = []
    for ym, sub in df.groupby("__ym", sort=True):
        sub = sub.drop(columns="__ym").reset_index(drop=True)
        groups.append((ym, sub))
    return groups


def split_by_month_two(main_df, phenom_df):
    """
    main_df と phenom_df を、同じ「年-月」の境界で同時に分割する。
    [(年月文字列, main部分df, phenom部分df), ...] を返す。
    """
    m = main_df.copy()
    p = phenom_df.copy()
    m["__ym"] = m["datetime"].dt.strftime("%Y-%m")
    p["__ym"] = p["datetime"].dt.strftime("%Y-%m")
    groups = []
    for ym in sorted(set(m["__ym"]) | set(p["__ym"])):
        msub = m[m["__ym"] == ym].drop(columns="__ym").reset_index(drop=True)
        psub = p[p["__ym"] == ym].drop(columns="__ym").reset_index(drop=True)
        groups.append((ym, msub, psub))
    return groups


def _draw_lane_panel(ax2, phenom_df, phenom_cols, encoded):
    """下段パネル（J〜AP列33行レーン）の描画処理を共通化した関数。"""
    ptimes = phenom_df["datetime"]
    n_lanes = len(phenom_cols)

    for row, col_letter in enumerate(phenom_cols):
        col_values = encoded[col_letter].to_numpy()
        slash_mask = col_values == 0
        if slash_mask.any():
            ax2.vlines(ptimes[slash_mask], row - 0.35, row + 0.35,
                       color=SLASH_COLOR, linewidth=1.0, alpha=0.9, zorder=2)
        for code in range(1, 11):
            code_mask = col_values == code
            if code_mask.any():
                ax2.vlines(ptimes[code_mask], row - 0.42, row + 0.42,
                           color=PHENOM_COLORS[code], linewidth=2.4, zorder=3)

    for row in range(n_lanes + 1):
        ax2.axhline(row - 0.5, color="#ececec", linewidth=0.5, zorder=1)

    ax2.set_yticks(range(n_lanes))
    ax2.set_yticklabels([f"列{c}" for c in phenom_cols], fontsize=8)
    ax2.set_ylim(n_lanes - 0.5, -0.5)
    ax2.set_xlabel("日時")
    ax2.set_ylabel("J〜AP列（現象記録列）", fontsize=10)

    legend_handles = [Patch(facecolor=SLASH_COLOR, label="「/」現象なし")]
    for code in range(1, 11):
        legend_handles.append(Patch(facecolor=PHENOM_COLORS[code],
                                    label=f"{code}: {PHENOM_LABELS[code]}"))
    ax2.legend(handles=legend_handles, bbox_to_anchor=(1.0, 1.0), loc="upper left",
               fontsize=8, borderaxespad=0., title="現象コード", title_fontsize=9)


def _finalize_figure(fig, ax1, ax2, times, fig_width):
    """時刻軸の共通設定と保存前の最終処理"""
    ax2.xaxis_date()
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
    n_ticks_target = max(10, int(fig_width / 1.3))
    ax2.xaxis.set_major_locator(
        mdates.AutoDateLocator(minticks=n_ticks_target, maxticks=n_ticks_target * 2)
    )
    ax1.set_xlim(times.iloc[0], times.iloc[-1])
    fig.autofmt_xdate()


def _make_fig(times, phenom_cols):
    """共通のfigure・2段subplotsを作成して返す"""
    n_lanes = len(phenom_cols)
    fig_width = compute_fig_width(times)
    lane_panel_height = max(6.0, n_lanes * 0.22)
    fig_height = 5.5 + lane_panel_height
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=(fig_width, fig_height), sharex=True,
        gridspec_kw={"height_ratios": [5.5, lane_panel_height], "hspace": 0.04},
    )
    return fig, ax1, ax2, fig_width


def plot_temp_humid_dew(main_df, phenom_df, phenom_cols, location_name, out_path):
    """
    グラフ①: 気温・露点温度・相対湿度 の複合グラフ + J〜AP列レーン

      左軸（℃）: 気温(℃) — 赤  /  露点温度(℃) — ティール
      右軸（％）: 相対湿度(％) — 紫（半透明）
    """
    times   = main_df["datetime"]
    encoded = phenom_df[phenom_cols].map(encode_phenomena_cell)
    fig, ax1, ax2, fig_width = _make_fig(times, phenom_cols)

    # 左軸: 気温・露点温度（℃）
    l1, = ax1.plot(times, main_df["気温(℃)"],    color="#e74c3c", linewidth=1.1,
                   label="気温(℃)", zorder=3)
    l2, = ax1.plot(times, main_df["露点温度(℃)"], color="#16a085", linewidth=1.1,
                   label="露点温度(℃)", zorder=3)
    ax1.set_ylabel("気温・露点温度（℃）", fontsize=10)
    ax1.grid(True, alpha=0.25)

    # 右軸: 相対湿度（%）
    ax1r = ax1.twinx()
    l3, = ax1r.plot(times, main_df["相対湿度(％)"], color="#8e44ad", linewidth=0.9,
                    alpha=0.65, label="相対湿度(％)", zorder=2)
    ax1r.set_ylabel("相対湿度（％）", fontsize=10)
    ax1r.set_ylim(0, 115)

    ax1.set_title(
        f"【{location_name}】気温・露点温度（左軸℃）・相対湿度（右軸%） と J〜AP列（現象コード）の関係",
        fontsize=13)
    ax1.legend(handles=[l1, l2, l3], loc="upper left", fontsize=9)

    _draw_lane_panel(ax2, phenom_df, phenom_cols, encoded)
    _finalize_figure(fig, ax1, ax2, times, fig_width)

    fig.savefig(out_path, dpi=SAVE_DPI, bbox_inches="tight")
    plt.close(fig)
    print(f"保存しました: {out_path}（横幅 約{fig_width:.0f}インチ）")


def plot_wind_precip(main_df, phenom_df, phenom_cols, location_name, out_path):
    """
    グラフ②: 風速・降水量 の複合グラフ + J〜AP列レーン

      左軸（m/s）: 風速(m/s) — 緑の実線
      右軸（mm） : 降水量(mm) — 青の棒グラフ（半透明）
    """
    times   = main_df["datetime"]
    encoded = phenom_df[phenom_cols].map(encode_phenomena_cell)
    fig, ax1, ax2, fig_width = _make_fig(times, phenom_cols)

    # 左軸: 風速（m/s）
    l1, = ax1.plot(times, main_df["風速(m/s)"], color="#27ae60", linewidth=1.1,
                   label="風速(m/s)", zorder=3)
    ax1.set_ylabel("風速（m/s）", fontsize=10)
    ax1.grid(True, alpha=0.25)

    # 右軸: 降水量（mm） — 棒グラフ
    ax1r = ax1.twinx()
    ax1r.bar(times, main_df["降水量(mm)"], width=0.03, color="#2980b9",
             alpha=0.55, label="降水量(mm)", zorder=2)
    pmax = main_df["降水量(mm)"].dropna().max() if not main_df["降水量(mm)"].dropna().empty else 1.0
    ax1r.set_ylim(0, max(pmax * 3.5, 2.0))
    ax1r.set_ylabel("降水量（mm）", fontsize=10)

    ax1.set_title(
        f"【{location_name}】風速（左軸m/s）・降水量（右軸mm） と J〜AP列（現象コード）の関係",
        fontsize=13)
    ax1.legend(handles=[l1, matplotlib.patches.Patch(color="#2980b9", alpha=0.55,
               label="降水量(mm)")], loc="upper left", fontsize=9)

    _draw_lane_panel(ax2, phenom_df, phenom_cols, encoded)
    _finalize_figure(fig, ax1, ax2, times, fig_width)

    fig.savefig(out_path, dpi=SAVE_DPI, bbox_inches="tight")
    plt.close(fig)
    print(f"保存しました: {out_path}（横幅 約{fig_width:.0f}インチ）")


def plot_combo_by_month(main_df, phenom_df, phenom_cols, location_name, out_dir):
    """2種類の複合グラフを月ごとに分割して出力するラッパー関数"""
    for ym, msub, psub in split_by_month_two(main_df, phenom_df):
        if msub["datetime"].nunique() < 2:
            continue
        loc = f"{location_name}（{ym}）"
        plot_temp_humid_dew(
            msub, psub, phenom_cols, loc,
            os.path.join(out_dir, f"{location_name}_①気温・湿度・露点×現象コード_{ym}.png"),
        )
        plot_wind_precip(
            msub, psub, phenom_cols, loc,
            os.path.join(out_dir, f"{location_name}_②風速・降水量×現象コード_{ym}.png"),
        )


# ---------------------------------------------------------------------------
# 5. メイン処理
# ---------------------------------------------------------------------------

# ----- Jupyter / Colab で直接セルに貼って実行する場合はここを書き換えてください -----
DEFAULT_INPUT_FILE = "只見_2026年1_5月_気象データ.xlsx"
DEFAULT_OUTPUT_DIR = "."
# -----------------------------------------------------------------------------


def _is_running_in_notebook():
    """Jupyter/ipykernel/Google Colab上で実行されているかどうかを判定する"""
    return (
        "ipykernel" in sys.modules
        or "IPython" in sys.modules
        or "google.colab" in sys.modules
    )


def _parse_args():
    """
    コマンドライン引数を解析する。Jupyter/Colabで直接セル実行された場合、
    sys.argvにカーネルの起動オプションが混ざることがあるため、
    .xlsx/.xls/フォルダ以外の引数は無視し、必要ならDEFAULT値にフォールバックする。
    """
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    args = [a for a in args if a.lower().endswith((".xlsx", ".xls")) or os.path.isdir(a)]

    filepath = None
    out_dir = None
    for a in args:
        if a.lower().endswith((".xlsx", ".xls")):
            filepath = a
        else:
            out_dir = a

    if filepath is None:
        if _is_running_in_notebook():
            print("（Jupyter/Colab上での実行を検出しました。DEFAULT_INPUT_FILE を使用します）")
            filepath = DEFAULT_INPUT_FILE
        else:
            print("使い方: python3 weather_visualizer.py 入力ファイル.xlsx [出力フォルダ]")
            sys.exit(1)

    if out_dir is None:
        out_dir = DEFAULT_OUTPUT_DIR

    return filepath, out_dir


def _try_colab_upload():
    """Google Colab上でファイルアップロード画面を表示し、選択されたxlsxパスを返す"""
    if "google.colab" not in sys.modules:
        return None
    try:
        from google.colab import files as colab_files
    except ImportError:
        return None

    print("入力ファイルが見つからなかったため、アップロード画面を表示します。")
    print("気象データのxlsxファイルを選択してください…")
    uploaded = colab_files.upload()
    for name in uploaded.keys():
        if name.lower().endswith((".xlsx", ".xls")):
            return os.path.abspath(name)
    return None


def _resolve_input_file(filepath):
    """ファイルが見つからない場合、Colabならアップロード、それ以外はinput()で確認する"""
    if os.path.isfile(filepath):
        return filepath

    print(f"入力ファイルが見つかりません: {filepath}")

    if "google.colab" in sys.modules:
        uploaded_path = _try_colab_upload()
        if uploaded_path and os.path.isfile(uploaded_path):
            return uploaded_path
        return None

    if _is_running_in_notebook():
        try:
            entered = input("入力ファイルのパスを入力してください（Enterでキャンセル): ").strip()
        except Exception:
            entered = ""
        if entered and os.path.isfile(entered):
            return entered
        return None

    return None


def main():
    filepath, out_dir = _parse_args()

    filepath = _resolve_input_file(filepath)
    if filepath is None:
        print("入力ファイルを取得できませんでした。DEFAULT_INPUT_FILE の値、")
        print("または実行時の引数（python3 weather_visualizer.py ファイル.xlsx）を確認してください。")
        sys.exit(1)

    if os.path.exists(out_dir) and not os.path.isdir(out_dir):
        raise FileExistsError(f"出力フォルダとして指定されたパスがファイルとして既に存在します: {out_dir}")
    os.makedirs(out_dir, exist_ok=True)

    main_df, phenom_df, phenom_cols = load_weather_data(filepath)

    # ファイル名から地点名を推測（例: 只見_2026年1_5月_気象データ.xlsx -> 只見）
    base = os.path.splitext(os.path.basename(filepath))[0]
    location_name = base.split("_")[0] if "_" in base else base

    # ① 気温・露点温度・相対湿度  ②風速・降水量 を月ごとに生成
    plot_combo_by_month(main_df, phenom_df, phenom_cols, location_name, out_dir)

    print("すべてのグラフを生成しました（月ごとに分割、各2種×月数枚ずつ出力）。")


if __name__ == "__main__":
    main()