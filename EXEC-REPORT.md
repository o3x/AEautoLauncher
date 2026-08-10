# EXEC-REPORT: AEselector（AEautoLauncher）リファクタリング実行

Mon Aug 10 19:35:48 JST 2026 望
実行環境: Windows / MSBuild 17.14.40（Visual Studio 2022 Community、.NET Framework 4.7.2）

## 事前対応: リポジトリの未同期状態

着手前に `git status` で未コミットの `CLAUDE.md` 差分と、originに1コミット分の遅れを検出した。差分内容を確認したところ、ローカルの未コミット差分は既にoriginにpush済みのコミット `3e89b1f`（大山さん自身によるもの、2026-07-04「AssemblyFileVersionの説明を追記」）と完全一致していた。単なるpull漏れと判断し、ローカル差分を破棄（`git checkout -- CLAUDE.md`）してから `git pull` で同期し、クリーンな状態から作業を開始した。

## 実施項目

| 項目 | 内容 | コミット |
|---|---|---|
| 項目0 | `refactor/2026-07`ブランチ作成（同期後クリーンだったため退避コミットは不要）。MSBuildベースラインビルド確認（0エラー・0警告） | (ブランチ作成のみ) |
| G1 | 追跡済みの`.vs/`・`bin/`・`AEautoLauncher.csproj.user`をgit追跡から除外（`obj/`は元々未追跡） | `1da684a` |
| C1 | `ResolveAePath`の`"UnKnown"`センチネルを`null`に置き換え | `aca5af5` |
| C2 | 3箇所重複していたバージョンビット抽出を`ExtractVersionBits`ヘルパーに集約 | `40d4012` |
| C3 | ヘッダーオフセット直値（48, 0x18, 0x24, 0x14, 0x27, 0x17, 0x25, 0x40, 0x68）を名前付き定数化 | `bd4221e` |
| C4 | `Program.cs`/`AssemblyInfo.cs`を0.4.6.0に更新、CHANGELOG.md追記 | `4a65c60` |

**`main`/`master`へのマージ・pushは行っていない**（計画書のやらないことリスト9番「push もしない（ローカルブランチまで）」に従い、`refactor/2026-07`ブランチに留めている）。

## 完了条件の実測結果

- 各項目コミット前に `MSBuild.exe AEautoLauncher.sln -p:Configuration=Release` を実行し、全項目で「ビルドに成功しました。0個の警告 0エラー」を確認
- G1: `git ls-files | grep -E "^\.vs/|^bin/|^obj/|\.user$"` がヒット0。ディスク上の`bin/Release/AEautoLauncher.exe`が無事に残っていることを確認
- C1: `grep -c "UnKnown" Program.cs` が0
- C2: `ExtractVersionBits`への置き換え前後で、`bytes[offset]`/`bytes[offset+1]`/`bytes[offset+2]`が元の`bytes[0x18]`/`[0x19]`/`[0x1A]`（CS5以前）、`[0x24]`/`[0x25]`/`[0x26]`（CS6以降ファイル版）、`[0x14]`/`[0x15]`/`[0x16]`（ホスト版）と1対1で対応することをコード上で突き合わせて確認
- C3: 定数への置き換え後も`OffsetFileVersion + 3 = 0x27`（revision）、`OffsetHostVersion + 3 = 0x17`、`OffsetFileVersion + 1 = 0x25`（Mac判定）が元の直値と一致することを確認。残存する直値は定数定義自身のみ
- C4: ビルド後のEXEを引数なしで起動し、`Start-Process`＋`MainWindowTitle`取得で確認したところ、ダイアログタイトルが `AEautoLauncher Version 0.4.6.0` であることを確認（起動後即座に`Stop-Process`で終了）

### 0-c 手動特性確認について

計画書の0-cは「Ctrlキーを押しながら実行」を要求するが、非対話環境（スクリプトからの起動）では物理的なキー押下状態を再現できないため、**実施不可**と判断した。計画書が想定するフォールバック（「実施不可の場合はビルド成功のみを完了条件とする」）に従い、C2・C3では計画書の代替指示どおり「置き換え前後の演算式をコード上で1対1に突き合わせて確認」する方法で検証した。

## 発見事項（計画書に無い問題）

なし。

## 未実施項目

なし（項目0・G1・C1〜C4すべて完了）。

## 大山さんへの確認事項

1. **`.aep`ファイルでの実機動作確認をお願いします**（Ctrlキー押下時のデバッグ表示、通常起動でのAE自動起動）。特にC2・C3はAE非公開仕様のビット演算・オフセットを扱うため、実ファイルでの動作確認が望ましいです
2. `main`へのマージ方針をご指示ください（LinkOpenerと同様、計画書に明記がないため`refactor/2026-07`ブランチのまま止めています）

## 詰まった点

- Git Bashから`msbuild`を`/p:`形式で呼ぶとパス展開されてしまうため、`-p:`形式に切り替えて対応した
- vswhereでのMSBuildパス検索も同様にGit Bashのパス展開の影響を受けたため、Git Bash形式パス（`/c/Program Files (x86)/...`）で実行した
