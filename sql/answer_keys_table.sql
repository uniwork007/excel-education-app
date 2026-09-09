-- =====================================================
-- answer_keys テーブル定義（最新版）
-- ステージ0〜8 の正解データ・型チェックデータを管理する
-- =====================================================

create table answer_keys (
  id             bigint generated always as identity primary key,
  stage_id       integer not null,
  cell           text not null,

  -- 値チェック用（ステージ0など）：期待される入力値
  -- 全角/半角・前後空白は正規化して比較するため、素の値を登録してOK
  expected_value text,

  -- 型チェック用（ステージ1など）：期待されるデータ型
  --   'n' = 数値（直接入力した数字）
  --   's' = 文字列（アポストロフィ付きや文字列として格納されたもの）
  --   'd' = 日付・時刻
  -- expected_value と expected_type は片方だけでも両方でも登録可
  expected_type  text check (expected_type in ('n', 's', 'd') or expected_type is null),

  -- 受講生が間違えた場合に表示するヒントメッセージ
  hint           text,

  created_at     timestamptz not null default now()
);

-- 同じステージ内で同じセルを重複登録できないようにする
create unique index answer_keys_stage_cell_idx on answer_keys (stage_id, cell);