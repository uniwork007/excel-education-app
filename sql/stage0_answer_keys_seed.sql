-- stage0_template.xlsx の採点対象セル（D8, D9）に対応する正解データ
-- answer_keys テーブル作成後、実際の課題ファイルの配置に合わせて実行してください

insert into answer_keys (stage_id, cell, expected_value, hint) values
  (0, 'D8', '山田太郎', null),
  (0, 'D9', '1500', '半角数字で入力しましょう');

-- D10（ひとことメモ）、D14〜D16（観察コーナー）はあえて登録していません。
-- 採点対象外のセルとして扱われます。
