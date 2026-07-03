# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

`.aep` ファイルのバイナリヘッダーを解析して After Effects のバージョンを判別し、対応する AE を自動起動する Windows 向けランチャー。フォームを持たないコンソールアプリ（WinExe）として動作し、`.aep` の関連付けに設定して使う。

## ビルド

- **ターゲット**: .NET Framework 4.7.2 / x64（Release）
- Visual Studio または MSBuild を使用

```
msbuild AEautoLauncher.sln /p:Configuration=Release
```

成果物: `bin\Release\AEautoLauncher.exe`

## アーキテクチャ

`Program.cs` 1ファイルのみで完結。主な処理フロー：

```
Main()
 └─ ExecuteLauncher()
     ├─ GetAeVersionFromFile()   // AEPバイナリヘッダー解析 → バージョン番号（整数）
     ├─ ResolveAePath()          // バージョン番号 → AE 実行ファイルパス
     │   └─ TryResolvePath()     // 複数フォルダ候補から存在するものを選択
     ├─ FindLatestInstalledAE()  // フォールバック: インストール済み最新 AE を検索
     └─ LaunchAfterEffects()     // AE 起動（Ctrl 押下中はデバッグ表示のみ）
```

### AEP ヘッダー解析の仕様

- マジックナンバー: `RIFF`/`RIFX` + offset 8 に `Egg!`
- CS5 以前: offset `0x18` のバイト列からバージョンをビット演算で抽出
- CS6 以降: offset `0x18 == 0x68` で判別し、offset `0x24` からバージョン抽出
- バージョン番号（整数）と年号の対応: v14〜v16 → `2003 + version` 年、v17〜v21 → 同式、v22 以降 → `2000 + version` 年

### バージョン → パスのマッピング

`ResolveAePath()` がバージョン整数をキーにフォルダ名を決定。  
未知バージョンやパスが存在しない場合は `FindLatestInstalledAE()` でインストール済み最新版にフォールバックし、ユーザーに確認ダイアログを表示。

## 動作モード

- **通常モード**: AE を起動してすぐ終了
- **デバッグモード**: 起動時に `Ctrl` キーを押し続けると AE を起動せず検出バージョン情報をダイアログ表示

## バージョン管理・リリースフロー

1. `Program.cs` 冒頭の `Version` と `Updated` を更新
2. `Properties\AssemblyInfo.cs` の `AssemblyVersion` / `AssemblyFileVersion` を更新  
   （`AssemblyFileVersion` が `Application.ProductVersion` の値となり、全ダイアログのタイトルバーに表示される）
3. CHANGELOG.md を更新（日時形式: `Sun Jan 12 12:44:00 JST 2026`）
4. Release ビルド → `gh release create vX.Y.Z` で GitHub Release を作成し EXE を添付
