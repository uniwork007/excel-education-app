import os
import re
import zipfile
import unicodedata
from flask import Flask, request, jsonify, render_template, redirect, url_for
import openpyxl
from dotenv import load_dotenv
from supabase import create_client, Client

# .envファイルから環境変数を読み込む
load_dotenv()

app = Flask(__name__)

# --- Supabase接続設定とURL補正ロジック ---
raw_url = os.environ.get("SUPABASE_URL", "").strip()
if "/rest/v1" in raw_url:
  raw_url = raw_url.split("/rest/v1")[0]

SUPABASE_URL = raw_url.rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "").strip()

print("--- 【修正後】起動時に読み込んだURL確認 ---")
print(f"URL: [{SUPABASE_URL}]")
print("-------------------------------------------")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# 各ステージのタイトル定義
STAGE_TITLES = {
    0: "基本操作",
    1: "データ型",
    2: "基本関数",
    3: "参照の理解",
    4: "応用関数",
    5: "データ整形",
    6: "可視化",
    7: "分析基礎",
    8: "総合演婚"
}


def normalize_for_compare(value):
  """初心者の入力揺れ（全角/半角、前後の空白、数値/文字列の違いなど）を
  吸収したうえで比較できるように正規化する。
  """
  if value is None:
    return ""
  # 整数として扱えるfloatは末尾の.0を除去して比較する（1500.0 と 1500 を同一視）
  if isinstance(value, float) and value.is_integer():
    value = int(value)
  text = unicodedata.normalize('NFKC', str(value))
  return text.strip()


def analyze_excel_file(file_stream, stage_id):
  """Excelファイルを解析して、躓き要素や数式をリストで返す"""
  wb = openpyxl.load_workbook(file_stream, data_only=False)

  # stage4は2シート構成。「マスタ」以外のシートを採点対象とする
  if stage_id == 4 and len(wb.sheetnames) > 1:
    ws = next((wb[s] for s in wb.sheetnames if s != "マスタ"), wb.active)
  else:
    ws = wb.active
  detected_errors = []
  functions_used = set()  # ステージ8で使用関数を収集するための集合

  # --- チェック対象セルの取得 ---
  # answer_keys に登録されているセル座標をチェック対象とする。
  # 登録がない場合は全セルをスキャン（既存挙動を維持）。
  # これにより説明文セルへの誤検知を防ぐ。
  try:
    ak_response = supabase.table("answer_keys").select("cell").eq(
        "stage_id", stage_id).execute()
    target_cells = {
        row["cell"].upper() for row in (ak_response.data or []) if row.get("cell")
    }
  except Exception:
    target_cells = set()

  def should_check(cell):
    """チェック対象セルかどうかを判定する。
    answer_keys に登録がある場合はそのセルのみ。
    登録がない場合は全セルをスキャン。
    """
    if not target_cells:
      return True
    return cell.coordinate.upper() in target_cells

  for row in ws.iter_rows(values_only=False):
    for cell in row:
      val = str(cell.value) if cell.value else ""

      # --- 共通のチェック（チェック対象セルのみ） ---
      if not should_check(cell):
        continue

      if any(err in val for err in ["#REF!", "#VALUE!", "#NAME?", "#DIV/0!"]):
        if "#NAME?" not in val:
          detected_errors.append(
              f"セル {cell.coordinate}: 数式エラー（{val}）が発生しています。")
          continue

      # 全角文字のチェックは「数式として入力されたセル」のみを対象とする。
      # 説明文セルの「（采点対象外）」などの普通の日本語テキストまで誤検知しないため。
      looks_like_formula = val.startswith("=") or val.startswith("＝")
      if looks_like_formula and ("＝" in val or "（" in val or "）" in val):
        detected_errors.append(
            f"セル {cell.coordinate}: ⚠️数式に全角文字（＝やかっこ）が混入しています。")
        continue

      # --- ステージ1：データ型のチェック ---
      if stage_id == 1:
        if "円" in val and not val.startswith("="):
          detected_errors.append(
              f"セル {cell.coordinate}: ⚠️値に「円」が直接入力されています。数値の後ろに単位をつけたい場合は『セルの書式設定』を使いましょう。"
          )
          continue

      # --- ステージ2：基本関数のチェック ---
      if stage_id == 2:
        upper_val = val.upper().replace(" ", "")
        if "#NAME?" in val or "AVARAGE" in upper_val or "SAM(" in upper_val:
          if "AVARAGE" in upper_val or "AVE(" in upper_val:
            detected_errors.append(
                f"セル {cell.coordinate}: ⚠️AVERAGE関数のスペルミス（AVARAGEやAVEなど）の疑いがあります。"
            )
          elif "SAM(" in upper_val:
            detected_errors.append(
                f"セル {cell.coordinate}: ⚠️SUM関数のスペルミス（SAM）の疑いがあります。")
          else:
            detected_errors.append(
                f"セル {cell.coordinate}: ⚠️関数名が間違っているため、#NAME? エラーが発生しています。")
          continue

        if (upper_val.startswith("=SUM(") or upper_val.startswith("=AVERAGE(")
            ) and "," in upper_val and ":" not in upper_val:
          detected_errors.append(
              f"セル {cell.coordinate}: ⚠️関数の範囲がコロン（:）ではなくカンマ（,）で区切られているため、2つのセルしか計算されていません。"
          )

        if upper_val.startswith("=IF(") and upper_val.count(",") == 1:
          detected_errors.append(
              f"セル {cell.coordinate}: ⚠️IF関数の引数が足りません。条件に合わない（偽の）場合の表示内容も設定しましょう。"
          )

      # --- ステージ3：参照の理解（数式抽出＋絶対参照チェック） ---
      if stage_id == 3:
        if val.startswith("="):
          detected_errors.append(
              f"【数式確認】セル {cell.coordinate}: {val}")

          # 全角文字の混入チェック（致命的バグのみ警告）
          if "＝" in val or "（" in val or "）" in val:
            detected_errors.append(
                f"⚠️ セル {cell.coordinate}: 数式に全角文字が混入しているため、計算されていません。")

          # F列のデータ行（10〜14行）は消費税率D4への参照が必須。
          # 「$D$4」が正解だが、「D$4」（行のみ絶対）も下コピーでズレないので許容。
          # 「$D4」（列のみ絶対）は下コピーで行番号がズレるため不正解として検出する。
          if cell.column == 6 and 10 <= cell.row <= 14:
            upper_val = val.upper().replace(" ", "")
            # 正解：$D$4 または D$4（行固定があればコピーしてもD4を参照し続ける）
            ok = "$D$4" in upper_val or "D$4" in upper_val
            if not ok:
              if "$D4" in upper_val:
                # 列のみ絶対参照：列は固定されているが行がズレる
                detected_errors.append(
                    f"セル {cell.coordinate}: ⚠️「$D4」になっています。"
                    f"下にコピーすると行番号がズレて D5, D6...と参照がずれてしまいます。"
                    f"「$D$4」と列・行の両方に$を付けましょう（F4キーを2回押すと$D4、もう1回で$D$4になります）。")
              else:
                detected_errors.append(
                    f"セル {cell.coordinate}: ⚠️消費税率のセル（D4）が絶対参照（$D$4）になっていません。"
                    f"数式をコピーするとズレてしまいます。F4キーで$マークを付けましょう。")

      # --- ステージ4：応用関数のチェック ---
      if stage_id == 4:
        upper_val = val.upper().replace(" ", "")

        # 数式が入っているセルのみを対象にする
        if upper_val.startswith("="):

          # ① VLOOKUP関数の徹底チェック
          if "VLOOKUP(" in upper_val:
            # 1. 完全一致（第4引数）の指定忘れチェック
            # VLOOKUPは引数4つが正しい形式。カンマが3つ未満＝引数が足りない。
            # 末尾が「,FALSE)」または「,0)」かどうかで判定する。
            vlookup_args = upper_val.count(",")
            has_false = upper_val.endswith(",FALSE)") or upper_val.endswith(",0)")
            if vlookup_args < 3 or not has_false:
              detected_errors.append(
                  f"セル {cell.coordinate}: ⚠️VLOOKUP関数の第4引数（検索方法）に 'FALSE' または '0' が指定されていません。"
                  f"これがないと、完全に一致するデータではなく『一番近いデータ』を返すことがあり、実務で大事故の原因になります。")

            # 2. マスタ範囲の絶対参照チェック
            # 「$A$3」形式（完全絶対）が正解。
            # 「$A3」（列のみ）は下コピーで行がズレるため不正解として検出する。
            import re as _re
            # マスタ!$A$3:$C$10 のような完全絶対参照パターンを探す
            has_full_abs = bool(_re.search(r'\$[A-Z]+\$\d+', upper_val))
            has_col_only = bool(_re.search(r'\$[A-Z]+\d+', upper_val)) and not has_full_abs
            if has_col_only:
              detected_errors.append(
                  f"セル {cell.coordinate}: ⚠️マスタ範囲の行番号に$がついていません（例：$A3）。"
                  f"数式を下にコピーするとマスタ範囲の行がズレてしまいます。"
                  f"「$A$3:$C$10」のように列・行の両方に$を付けましょう。")
            elif not has_full_abs:
              detected_errors.append(
                  f"セル {cell.coordinate}: ⚠️VLOOKUPのマスタ範囲に絶対参照（$）がついていません。"
                  f"数式をコピーするとマスタ範囲がズレてしまいます。"
                  f"「マスタ!$A$3:$C$10」のように$マークを付けましょう。")

          # ② XLOOKUP関数のチェック（最新の表計算ソフト対応）
          elif "XLOOKUP(" in upper_val:
            # XLOOKUPは最低3つの引数（検索値, 検索範囲, 戻り範囲）が必要です。
            # カンマが最低2つ以上含まれているかチェックします
            comma_count = upper_val.count(",")
            if comma_count < 2:
              detected_errors.append(
                  f"セル {cell.coordinate}: ⚠️XLOOKUP関数の引数が足りません（最低3つの指定が必要です）。 "
                  f"『何を探すか』『どこを探すか』『どこから結果を戻すか』が正しく区切られているか確認しましょう。")

            # 検索範囲や戻り範囲の絶対参照忘れ
            if not "$" in upper_val:
              detected_errors.append(
                  f"セル {cell.coordinate}: ⚠️XLOOKUP関数の検索範囲・戻り範囲に絶対参照（$）が使われていない可能性があります。"
              )

      # --- ステージ8：総合演習（使用されている関数名をセル単位で収集） ---
      if stage_id == 8:
        if val.startswith("="):
          upper_val = val.upper().replace(" ", "")
          functions_used.update(re.findall(r'([A-Z]+)\(', upper_val))

  # --- ステージ2：関数使用チェック（ループ外・シート単位） ---
  # スペルミス・範囲ミスはループ内で検出済み。
  # ここでは「指定された関数が実際に使われているか」を採点対象セルごとに確認する。
  if stage_id == 2:
    # {セル座標: (期待する関数名, ラベル)} の対応表
    REQUIRED_FUNCTIONS = {
        "F8":  ("SUM",     "田中 一郎の合計"),
        "F9":  ("SUM",     "鈴木 花子の合計"),
        "F10": ("SUM",     "佐藤 次郎の合計"),
        "F11": ("SUM",     "山田 三枝の合計"),
        "F12": ("SUM",     "伊藤 四朗の合計"),
        "F15": ("COUNT",   "受験者数"),
        "F16": ("AVERAGE", "平均点"),
        "F17": ("MAX",     "最高点"),
        "F18": ("MIN",     "最低点"),
    }
    for cell_ref, (func_name, label) in REQUIRED_FUNCTIONS.items():
      try:
        cell_val = str(ws[cell_ref].value or "").upper().replace(" ", "")
      except Exception:
        continue

      if not cell_val.startswith("="):
        detected_errors.append(
            f"セル {cell_ref}（{label}）: ⚠️数式が入力されていません。"
            f"{func_name}関数を使って求めましょう。")
      elif f"{func_name}(" not in cell_val:
        detected_errors.append(
            f"セル {cell_ref}（{label}）: ⚠️{func_name}関数が使われていません。"
            f"=で始まる数式の中に {func_name}( ) を使って求めましょう。")

  # --- ステージ4：VLOOKUP使用チェック（ループ外・シート単位） ---
  if stage_id == 4:
    # C列・D列（商品名・単価）にVLOOKUPが使われているかを確認
    VLOOKUP_REQUIRED = {
        "C9":  "商品名（C列）", "C10": "商品名（C列）", "C11": "商品名（C列）",
        "C12": "商品名（C列）", "C13": "商品名（C列）", "C14": "商品名（C列）",
        "D9":  "単価（D列）",   "D10": "単価（D列）",   "D11": "単価（D列）",
        "D12": "単価（D列）",   "D13": "単価（D列）",   "D14": "単価（D列）",
    }
    # シート名「課題4」を探す（受講生がシート名を変えていた場合も考慮）
    target_ws = None
    for sheet_name in wb.sheetnames:
      if sheet_name != "マスタ":
        target_ws = wb[sheet_name]
        break

    if target_ws:
      for cell_ref, label in VLOOKUP_REQUIRED.items():
        try:
          cell_val = str(target_ws[cell_ref].value or "").upper().replace(" ", "")
        except Exception:
          continue
        if not cell_val.startswith("="):
          detected_errors.append(
              f"セル {cell_ref}（{label}）: ⚠️数式が入力されていません。"
              f"VLOOKUP関数を使って商品マスタから自動取得しましょう。")
        elif "VLOOKUP(" not in cell_val:
          detected_errors.append(
              f"セル {cell_ref}（{label}）: ⚠️VLOOKUP関数が使われていません。"
              f"=VLOOKUP(検索値, マスタ!$A$3:$C$10, 列番号, FALSE) の形式で入力しましょう。")

  # --- ステージ5：データ整形のチェック（シート単位のためループの外で判定） ---
  if stage_id == 5:
    # ① テーブル化されているか
    if not ws.tables:
      detected_errors.append(
          "⚠️テーブル化（Ctrl+T）が行われていません。表全体を「テーブル」として登録しましょう。")

    # ② オートフィルタが設定されているか
    #    （テーブル化していれば自動的にフィルタも付くので、テーブル自身のautoFilterも確認）
    has_filter = bool(ws.auto_filter.ref) or any(
        table.autoFilter and table.autoFilter.ref
        for table in ws.tables.values())
    if not has_filter:
      detected_errors.append(
          "⚠️並べ替え・フィルタ（オートフィルタ）が設定されていません。")

    # ③ 条件付き書式が設定されているか
    has_conditional_formatting = any(
        cf_range.rules for cf_range in ws.conditional_formatting)
    if not has_conditional_formatting:
      detected_errors.append(
          "⚠️条件付き書式が設定されていません。異常値のハイライトなどを設定しましょう。")

  # --- ステージ6：可視化のチェック（グラフはopenpyxlの非公開APIに依存するため try/except で保護） ---
  if stage_id == 6:
    try:
      charts = ws._charts
    except AttributeError:
      charts = []
      detected_errors.append(
          "【参考情報】このバージョンのopenpyxlではグラフ情報を読み取れませんでした。目視で確認してください。")

    if not charts:
      detected_errors.append("⚠️グラフが検出されませんでした。")
    else:
      detected_errors.append(f"【参考情報】検出されたグラフ数: {len(charts)}個")

      for i, chart in enumerate(charts, start=1):
        chart_type = type(chart).__name__

        # タイトルのテキストを安全に取り出す
        try:
          title_text = chart.title.tx.rich.p[0].r[0].t
        except (AttributeError, IndexError, TypeError):
          title_text = None

        detected_errors.append(
            f"【参考情報】グラフ{i}: 種類={chart_type}, タイトル={title_text or '(未設定)'}"
        )

        if not title_text:
          detected_errors.append(f"⚠️グラフ{i}にタイトルが設定されていません。")

        # データ系列が1件も紐付いていない＝データ未選択のまま挿入された空グラフ
        if not chart.series:
          detected_errors.append(
              f"⚠️グラフ{i}にデータが紐付けられていません（空のグラフの可能性があります）。")

  # --- ステージ7：分析基礎のチェック（データ検証は自動判定、ピボットは存在確認のみ） ---
  if stage_id == 7:
    # ① データ検証（入力規則）が設定されているか
    if not ws.data_validations.dataValidation:
      detected_errors.append(
          "⚠️データ検証（入力規則）が設定されていません。「データ」タブの『データの入力規則』を確認しましょう。"
      )

    # ② ピボットテーブルが使われているか（openpyxlに直接読むAPIがないため、
    #    xlsxをzipとして開きxl/pivotCache配下の有無で存在確認する）
    if has_pivot_table(file_stream):
      detected_errors.append(
          "【参考情報】ピボットテーブルが検出されました。集計軸などの中身は目視で確認してください。")
    else:
      detected_errors.append(
          "⚠️ピボットテーブルが検出されませんでした。「挿入」タブの『ピボットテーブル』を使いましょう。"
      )

  # --- ステージ8：総合演習（使用機能のチェックリストを提示。合否判定は教員の目視に委ねる） ---
  if stage_id == 8:
    features_used = []

    if ws.tables:
      features_used.append("テーブル")

    has_cf = any(cf_range.rules for cf_range in ws.conditional_formatting)
    if has_cf:
      features_used.append("条件付き書式")

    if ws.data_validations.dataValidation:
      features_used.append("データ検証")

    if has_pivot_table(file_stream):
      features_used.append("ピボットテーブル")

    try:
      chart_count = len(ws._charts)
    except AttributeError:
      chart_count = 0
    if chart_count > 0:
      features_used.append(f"グラフ（{chart_count}個）")

    if functions_used:
      features_used.append(f"使用関数: {', '.join(sorted(functions_used))}")

    detected_errors.append(
        f"【検出機能一覧】{', '.join(features_used) if features_used else '検出された機能はありませんでした'}"
    )

  # --- ステージ0：基本操作のチェック（Supabaseの正解データと突き合わせる） ---
  if stage_id == 0:
    try:
      response = supabase.table("answer_keys").select("*").eq("stage_id",
                                                                0).execute()
      answer_keys = response.data or []
    except Exception as e:
      answer_keys = []
      detected_errors.append(
          f"【注意】正解データの取得に失敗しました（{e}）。管理者に確認してください。")

    if not answer_keys:
      detected_errors.append(
          "【注意】ステージ0の正解データがまだ登録されていません。管理者に確認してください。")
    else:
      for key in answer_keys:
        cell_ref = key.get("cell")
        expected_value = key.get("expected_value")
        hint = key.get("hint") or ""

        try:
          actual_value = ws[cell_ref].value
        except Exception:
          detected_errors.append(
              f"⚠️セル指定「{cell_ref}」が不正なため確認できませんでした。管理者に確認してください。")
          continue

        if normalize_for_compare(actual_value) != normalize_for_compare(
            expected_value):
          display_actual = actual_value if actual_value is not None else "(空欄)"
          hint_text = f" ヒント: {hint}" if hint else ""
          detected_errors.append(
              f"セル {cell_ref}: ⚠️期待される値と異なります（入力値: {display_actual}）。{hint_text}"
          )

  # --- ステージ1：データ型のチェック（openpyxlのdata_typeで数値/文字列/日付を判定） ---
  # 「1500」と入力しても「'1500」と入力しても見た目の文字列は同じになるため、
  # 値の比較ではなく cell.data_type（'n'=数値, 's'=文字列, 'd'=日付）で判定する。
  if stage_id == 1:
    type_labels = {'n': '数値', 's': '文字列', 'd': '日付', 'b': '真偽値', 'f': '数式'}

    try:
      response = supabase.table("answer_keys").select("*").eq(
          "stage_id", 1).execute()
      type_keys = [
          k for k in (response.data or []) if k.get("expected_type")
      ]
    except Exception as e:
      type_keys = []
      detected_errors.append(
          f"【注意】型チェック用データの取得に失敗しました（{e}）。管理者に確認してください。")

    for key in type_keys:
      cell_ref = key.get("cell")
      expected_type = key.get("expected_type")

      try:
        cell = ws[cell_ref]
      except Exception:
        detected_errors.append(
            f"⚠️セル指定「{cell_ref}」が不正なため確認できませんでした。管理者に確認してください。")
        continue

      actual_type = cell.data_type
      cell_val_str = str(cell.value) if cell.value is not None else ""

      # 日付はシリアル値として保存されdatatype='n'になる場合があるため
      # is_date_cell()で再判定し、expected_type=='d'なら正解扱いにする
      if expected_type == 'd' and is_date_cell(cell):
        continue

      if actual_type != expected_type:
        expected_label = type_labels.get(expected_type, expected_type)
        actual_label = type_labels.get(actual_type, actual_type)

        # 「円」付き入力（例:1000円）は数値欄への文字列保存として検出されるが、
        # セルループ内の「円チェック」で既に分かりやすいメッセージを出しているため
        # 初心者向けに型エラーメッセージは重複させない。
        if expected_type == 'n' and actual_type == 's' and "円" in cell_val_str:
          continue

        guidance = ""
        if expected_type == 'n' and actual_type == 's':
          guidance = "数字の前にアポストロフィ（\'）が付いていないか確認しましょう。"
        elif expected_type == 'd' and actual_type == 's':
          guidance = "「2026/9/8」のようにスラッシュ区切りで入力すると自動的に日付として認識されます。"

        detected_errors.append(
            f"セル {cell_ref}: ⚠️{expected_label}を入力してほしいところですが、"
            f"セルが左寄せ（文字列扱い）になっています。{guidance}")

  return detected_errors


def is_date_cell(cell):
  """openpyxlの data_type だけでなく、Excelのシリアル値範囲と書式も合わせて
  日付セルかどうかを判定する。
  Excelは日付を内部的にシリアル値（1900/1/1=1 起点の整数）として保存するため、
  number_format が 'General' のままでも日付として入力された場合がある。
  シリアル値の範囲：1（1900/1/1）〜 65380（2078/12/31）
  """
  if cell.data_type == 'd':
    return True
  fmt = (cell.number_format or "").lower()
  DATE_FORMAT_KEYWORDS = ["yy", "mm", "dd", "m/d", "d/m", "[$-", "年", "月", "日"]
  if any(k in fmt for k in DATE_FORMAT_KEYWORDS):
    return True
  val = cell.value
  if isinstance(val, (int, float)) and 1 <= val <= 65380:
    return True
  return False


def has_pivot_table(file_stream):
  """xlsxファイル内にピボットテーブルの定義が含まれているかを確認する。
  openpyxlはピボットテーブルを直接読み込めないため、xlsxをzipとして開き、
  xl/pivotCache配下にファイルが存在するかどうかで簡易的に判定する。
  """
  try:
    file_stream.seek(0)
    with zipfile.ZipFile(file_stream) as z:
      return any("pivotCache" in name for name in z.namelist())
  except (zipfile.BadZipFile, AttributeError):
    return False
  finally:
    file_stream.seek(0)


# --- 受講生用：ファイル提出画面 ---
@app.route('/')
def upload_page():
  return render_template('upload.html')


# --- 受講生用：ファイル受け取り・判定処理 ---
@app.route('/upload_progress', methods=['POST'])
def upload_progress():
  student_id = request.form.get('student_id')
  stage_id = int(request.form.get('stage_id'))
  file = request.files.get('excel_file')

  if not file or not file.filename.endswith('.xlsx'):
    return render_template('result.html',
                            success=False,
                            student_id=None,
                            stage_name=None,
                            status=None,
                            error_count=0,
                            errors=[],
                            info_logs=[],
                            message="有効な.xlsxファイルをアップロードしてください。"), 400

  try:
    try:
      supabase.table("students").insert({
          "student_id": student_id,
          "name": "テスト受講生"
      }).execute()
    except:
      pass

    # エラーおよび数式を解析
    all_logs = analyze_excel_file(file, stage_id)

    # 「⚠️」や「エラー」を含むものを警告としてカウント
    real_errors = [msg for msg in all_logs if "⚠️" in msg or "エラー" in msg]
    # 「【参考情報】」「【数式確認】」「【検出機能一覧】」は参考情報として分離
    info_logs = [msg for msg in all_logs if msg not in real_errors]
    error_count = len(real_errors)

    # ステータスの判定ルール（ステージ3・6・8は目視レビュー待ちへ）
    if stage_id == 3 or stage_id == 6 or stage_id == 8:
      status = "目視レビュー待ち"
    else:
      status = "要確認" if error_count > 0 else "提出済"

    progress_data = {
        "student_id": student_id,
        "stage_id": stage_id,
        "status": status,
        "error_logs": {
            "detected_errors": all_logs,
            "error_count": error_count
        }
    }
    supabase.table("progress_results").insert(progress_data).execute()

    return render_template('result.html',
                            success=True,
                            student_id=student_id,
                            stage_name=STAGE_TITLES.get(stage_id,
                                                         f"ステージ{stage_id}"),
                            status=status,
                            error_count=error_count,
                            errors=real_errors,
                            info_logs=info_logs,
                            message=None)
  except Exception as e:
    return render_template('result.html',
                            success=False,
                            student_id=student_id,
                            stage_name=None,
                            status=None,
                            error_count=0,
                            errors=[],
                            info_logs=[],
                            message=f"サーバーエラーが発生しました: {str(e)}"), 500


# --- 教員用：管理画面 ---
@app.route('/admin')
def admin_dashboard():
  response = supabase.table("progress_results").select("*").order(
      "updated_at", desc=True).execute()
  results = response.data
  for item in results:
    item['stage_name'] = STAGE_TITLES.get(item['stage_id'],
                                          f"ステージ{item['stage_id']}")
    error_logs = item.get('error_logs', {}) or {}
    item['errors'] = error_logs.get('detected_errors', [])
    item['error_count'] = error_logs.get('error_count', 0)
  return render_template('admin.html', results=results)


# --- 教員用：評価更新 ---
@app.route('/admin/review/<record_id>', methods=['POST'])
def review_stage(record_id):
  new_status = request.form.get('status')
  teacher_comment = request.form.get('human_review')
  supabase.table("progress_results").update({
      "status": new_status,
      "human_review": teacher_comment
  }).eq("id", record_id).execute()
  return redirect(url_for('admin_dashboard'))


# --- 教員用：ステージ0 正解データの一覧・登録画面 ---
@app.route('/admin/answer_keys')
def answer_keys_page():
  # 表示するステージをクエリパラメータで切り替えられるようにする（デフォルトはステージ0）
  stage_id = int(request.args.get('stage_id', 0))
  response = supabase.table("answer_keys").select("*").eq(
      "stage_id", stage_id).order("cell").execute()
  keys = response.data or []
  return render_template('answer_keys.html',
                          keys=keys,
                          stage_id=stage_id,
                          stage_titles=STAGE_TITLES)


# --- 教員用：正解データの追加 ---
@app.route('/admin/answer_keys/add', methods=['POST'])
def answer_keys_add():
  stage_id = int(request.form.get('stage_id'))
  cell = request.form.get('cell', '').strip().upper()
  expected_value = request.form.get('expected_value', '').strip()
  expected_type = request.form.get('expected_type', '').strip()
  hint = request.form.get('hint', '').strip()

  if not cell or (not expected_value and not expected_type):
    return "セルに加えて、期待値または期待する型のどちらかは必須です", 400

  try:
    supabase.table("answer_keys").insert({
        "stage_id": stage_id,
        "cell": cell,
        "expected_value": expected_value or None,
        "expected_type": expected_type or None,
        "hint": hint or None
    }).execute()
  except Exception as e:
    return f"登録に失敗しました（同じセルが既に登録されている可能性があります）: {str(e)}", 400

  return redirect(url_for('answer_keys_page', stage_id=stage_id))


# --- 教員用：正解データの削除 ---
@app.route('/admin/answer_keys/delete/<record_id>', methods=['POST'])
def answer_keys_delete(record_id):
  stage_id = request.form.get('stage_id', 0)
  supabase.table("answer_keys").delete().eq("id", record_id).execute()
  return redirect(url_for('answer_keys_page', stage_id=stage_id))


if __name__ == '__main__':
  # 社内LANの他のPCから接続できるように 0.0.0.0 をバインド
  app.run(host='0.0.0.0', port=5000, debug=True)
