-- ステージ0（および将来的な他ステージ）の正解データを管理するテーブル
create table answer_keys (
  id bigint generated always as identity primary key,
  stage_id integer not null,
  cell text not null,
  expected_value text not null,
  hint text,
  created_at timestamptz not null default now()
);

-- 同じステージ内で同じセルを重複登録できないようにする
create unique index answer_keys_stage_cell_idx on answer_keys (stage_id, cell);
