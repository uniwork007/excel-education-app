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
  ws = wb.active
  detected_errors = []
  functions_used = set()  # ステージ8で使用関数を収集するための集合

  for row in ws.iter_rows(values_only=False):
    for cell in row:
      val = str(cell.value) if cell.value else ""

      # --- 共通のチェック（全ステージ共通） ---
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

      # --- ステージ3：参照の理解（数式抽出の人間系判断モード） ---
      if stage_id == 3:
        if val.startswith("="):
          detected_errors.append(f"【数式確認】セル {cell.coordinate}: {val}")
          # 数式の中に全角が混ざっている致命的なバグだけ警告として残す
          if "＝" in val or "（" in val or "）" in val:
            detected_errors.append(
                f"⚠️ セル {cell.coordinate}: 数式に全角文字が混入しているため、計算されていません。")

      # --- ステージ4：応用関数のチェック ---
      if stage_id == 4:
        upper_val = val.upper().replace(" ", "")

        # 数式が入っているセルのみを対象にする
        if upper_val.startswith("="):

          # ① VLOOKUP関数の徹底チェック
          if "VLOOKUP(" in upper_val:
            # 1. 完全一致（第4引数）の指定忘れチェック
            # 引数が4つ未満、または4つ目（最後の引数）が 0/FALSE になっていない場合
            # カンマの数と末尾の記述を正規表現で厳密にチェックします
            if not (upper_val.endswith(",0)")
                    or upper_val.endswith(",FALSE)")
                    or upper_val.endswith(",0.0)")
                    or "FALSE," in upper_val or "0," in upper_val):
              detected_errors.append(
                  f"セル {cell.coordinate}: ⚠️VLOOKUP関数の第4引数（検索方法）に 'FALSE' または '0' が指定されていません。 "
                  f"これがないと、完全に一致するデータではなく『一番近いデータ』を勝手に探してしまい、実務で大事故の原因になります。"
              )

            # 2. 検索マスタ範囲（第2引数）の絶対参照忘れチェック
            # 数式を下にコピペしたときにマスタ範囲がズレるのを防ぐため
            # 数式全体に $ が含まれていない場合は警告
            if not "$" in upper_val:
              detected_errors.append(
                  f"セル {cell.coordinate}: ⚠️VLOOKUPの参照マスタ範囲に絶対参照（$）がついていない可能性があります。 "
                  f"数式を下にコピペした際、マスタの範囲まで一緒にズレてしまっていませんか？")

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

      if actual_type != expected_type:
        expected_label = type_labels.get(expected_type, expected_type)
        actual_label = type_labels.get(actual_type, actual_type)

        guidance = ""
        if expected_type == 'n' and actual_type == 's':
          guidance = (
              "セルが左寄せになっていませんか？"
              "数字の前にアポストロフィ（\'）が付いていないか確認しましょう。")
        elif expected_type == 'd' and actual_type == 's':
          guidance = (
              "セルが左寄せになっていませんか？"
              "日付は「2026/9/8」のように入力すると自動的に日付として認識されます。")

        detected_errors.append(
            f"セル {cell_ref}: ⚠️{expected_label}として入力してほしいところですが、"
            f"{actual_label}として保存されています。{guidance}")

  return detected_errors


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
    return "有効な.xlsxファイルをアップロードしてください", 400

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

    # 「⚠️」や「エラー」という文字が入っているものだけを、本当の「エラー」としてカウント
    real_errors = [msg for msg in all_logs if "⚠️" in msg or "エラー" in msg]
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

    return f"<h3>提出が完了しました！（自動検知されたエラー: {error_count}件）</h3><a href='/'>戻る</a>"
  except Exception as e:
    return f"エラー: {str(e)}", 500


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
