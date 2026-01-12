// Program.cs
// Version: 0.4.5.0
// Updated: Sun Jan 12 12:44:00 JST 2026

using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace AEautoLauncher
{
    static class Program
    {
        private const string ProgramFilesX86Adobe = @"C:\Program Files (x86)\Adobe\";
        private const string ProgramFilesX64Adobe = @"C:\Program Files\Adobe\";
        private const string AfterEffectsExePath = @"\Support Files\AfterFX.exe";

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ExecuteLauncher();
        }

        private static void ExecuteLauncher()
        {
            try
            {
                string[] args = Environment.GetCommandLineArgs();

                if (args.Length > 2)
                {
                    ShowMessage("複数のファイル選択には対応していません");
                    return;
                }

                if (args.Length != 2)
                {
                    ShowMessage("AE5.0以降に対応\rフォルダはデフォルト決め打ち\r拡張子AEPの関連づけをAEautoLauncherにしてください。");
                    return;
                }

                string aepPath = args[1];
                if (!File.Exists(aepPath))
                {
                    ShowMessage($"ファイルが見つかりません: {aepPath}");
                    return;
                }

                int aeVersion = GetAeVersionFromFile(aepPath, out string strVersionInfo);
                string aeInstallPath = ResolveAePath(aeVersion);

                if (aeInstallPath == "UnKnown" || !File.Exists(aeInstallPath))
                {
                    // インストール済みの最新バージョンを検出してフォールバック
                    string latestPath = FindLatestInstalledAE();
                    
                    if (latestPath == null)
                    {
                        ShowMessage($"After Effectsがインストールされていません。\r検出されたバージョン: {strVersionInfo}");
                        return;
                    }

                    string latestVersionName = Path.GetFileName(Path.GetDirectoryName(Path.GetDirectoryName(latestPath)));
                    DialogResult result = MessageBox.Show(
                        $"バージョン不明または未インストールのバージョンです。\r{latestVersionName}で起動しますか？\r検出されたバージョン: {strVersionInfo}",
                        $"AEautoLauncher Version {Application.ProductVersion}",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.OK)
                    {
                        LaunchAfterEffects(latestPath, aepPath, strVersionInfo);
                    }
                }
                else
                {
                    LaunchAfterEffects(aeInstallPath, aepPath, strVersionInfo);
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"エラーが発生しました: {ex.Message}");
            }
        }

        private static int GetAeVersionFromFile(string path, out string versionString)
        {
            int version = 0;
            versionString = "Unknown";
            
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (BinaryReader br = new BinaryReader(fs))
                {
                    byte[] bytes = br.ReadBytes(48); // ヘッダー読み込み
                    if (bytes.Length < 48) return 0;

                    // マジックナンバーチェック: RIFF/RIFX ... Egg!
                    // RIFF = 0x52, 0x49, 0x46, 0x46 (Little Endian)
                    // RIFX = 0x52, 0x49, 0x46, 0x58 (Big Endian)
                    // Egg! = 0x45, 0x67, 0x67, 0x21 (AEP識別子)
                    bool isRiff = (bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46);
                    bool isRifx = (bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x58);
                    bool isEgg  = (bytes[8] == 0x45 && bytes[9] == 0x67 && bytes[10] == 0x67 && bytes[11] == 0x21);

                    if ((!isRiff && !isRifx) || !isEgg)
                    {
                        return 0; // 有効なAEPファイルではない
                    }

                    bool isCs6OrLater = (bytes[0x18] == 0x68);

                    if (!isCs6OrLater)
                    {
                        // CS5以前のバージョン判定
                        version = ((bytes[0x18] << 1) & 0xF8) + ((bytes[0x19] >> 3) & 0x07);
                        int minor = ((bytes[0x19] << 1) & 0x0E) + (bytes[0x1A] >> 7);
                        int build = (bytes[0x1A] >> 3) & 0x0F;
                        versionString = $"{version}.{minor}.{build}";
                    }
                    else
                    {
                        // CS6以降のバージョン判定
                        version = ((bytes[0x24] << 1) & 0xF8) + ((bytes[0x25] >> 3) & 0x07);
                        int minor = ((bytes[0x25] << 1) & 0x0E) + (bytes[0x26] >> 7);
                        int build = (bytes[0x26] >> 3) & 0x0F;
                        int revision = bytes[0x27];
                        versionString = $"{version}.{minor}.{build}.{revision}";

                        // ホストバージョン情報を抽出（追加情報用）
                        int hostVer = ((bytes[0x14] << 1) & 0xF8) + ((bytes[0x15] >> 3) & 0x07);
                        int hostMinor = ((bytes[0x15] << 1) & 0x0E) + (bytes[0x16] >> 7);
                        int hostBuild = (bytes[0x16] >> 3) & 0x0F;
                        string hostVerString = $"{hostVer}.{hostMinor}.{hostBuild}.{bytes[0x17]}";

                        // プラットフォーム情報追加前にバージョン比較
                        bool versionsMatch = (versionString == hostVerString);

                        string platform = (bytes[0x25] & 0x40) == 0 ? "(Win)" : "(Mac)";
                        versionString += platform;

                        if (!versionsMatch)
                        {
                            versionString += $" [HostVersion:{hostVerString}]";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // ファイル読み込みエラー時はバージョン0を返す
                System.Diagnostics.Debug.WriteLine($"AEPファイル読み込みエラー: {ex.Message}");
            }

            return version;
        }

        /// <summary>
        /// AEバージョン番号からインストールパスを解決する
        /// </summary>
        private static string ResolveAePath(int version)
        {
            // 古いバージョン（32bit）のマッピング
            if (version == 5) return TryResolvePath(ProgramFilesX86Adobe, "After Effects 5.0", "After Effects 5.5");
            if (version == 6) return TryResolvePath(ProgramFilesX86Adobe, "After Effects 6.0", "After Effects 6.5");
            if (version == 7) return ProgramFilesX86Adobe + @"After Effects 7.0" + AfterEffectsExePath;
            if (version == 8) return ProgramFilesX86Adobe + @"Adobe After Effects CS3" + AfterEffectsExePath;
            if (version == 9) return ProgramFilesX86Adobe + @"Adobe After Effects CS4" + AfterEffectsExePath;
            
            // CS5 - CS6（64bit）
            if (version == 10) return TryResolvePath(ProgramFilesX64Adobe, "Adobe After Effects CS5", "Adobe After Effects CS5.5");
            if (version == 11) return ProgramFilesX64Adobe + @"Adobe After Effects CS6" + AfterEffectsExePath;
            
            // CCバージョン
            if (version == 12) return ProgramFilesX64Adobe + @"Adobe After Effects CC" + AfterEffectsExePath;
            if (version == 13) return TryResolvePath(ProgramFilesX64Adobe, "Adobe After Effects CC 2014", "Adobe After Effects CC 2015", "Adobe After Effects CC 2015.3");

            // CC 2017以降の自動マッピング (v14+)
            if (version >= 14 && version < 17)
            {
               return ProgramFilesX64Adobe + $@"Adobe After Effects CC {2003 + version}" + AfterEffectsExePath;
            }
            if (version >= 17 && version < 22)
            {
               return ProgramFilesX64Adobe + $@"Adobe After Effects {2003 + version}" + AfterEffectsExePath;
            }
            if (version >= 22)
            {
               return ProgramFilesX64Adobe + $@"Adobe After Effects {2000 + version}" + AfterEffectsExePath;
            }

            return "UnKnown";
        }

        /// <summary>
        /// 複数のフォルダ名候補から存在するパスを返す（優先度順）
        /// </summary>
        private static string TryResolvePath(string basePath, params string[] folderNames)
        {
            foreach (string folder in folderNames)
            {
                string path = basePath + folder + AfterEffectsExePath;
                if (File.Exists(path))
                {
                    return path;
                }
            }
            // どれも見つからない場合は最初の候補を返す（後のチェックでフォールバック処理へ）
            return basePath + folderNames[0] + AfterEffectsExePath;
        }

        /// <summary>
        /// インストール済みの最新After Effectsを検出して実行ファイルパスを返す
        /// </summary>
        private static string FindLatestInstalledAE()
        {
            string latestPath = null;
            int latestYear = 0;

            string[] searchPaths = { ProgramFilesX64Adobe, ProgramFilesX86Adobe };

            foreach (string basePath in searchPaths)
            {
                if (!Directory.Exists(basePath)) continue;

                foreach (string dir in Directory.GetDirectories(basePath, "*After Effects*"))
                {
                    string exePath = dir + AfterEffectsExePath;
                    if (!File.Exists(exePath)) continue;

                    string folderName = Path.GetFileName(dir);
                    int year = ExtractYearFromFolderName(folderName);

                    if (year > latestYear)
                    {
                        latestYear = year;
                        latestPath = exePath;
                    }
                }
            }

            return latestPath;
        }

        /// <summary>
        /// フォルダ名から年度を抽出（例: "Adobe After Effects 2024" → 2024）
        /// </summary>
        private static int ExtractYearFromFolderName(string folderName)
        {
            // "CC 2017", "CC 2018", "2020", "2024" などのパターンを検出
            Match match = Regex.Match(folderName, @"20\d{2}");
            
            if (match.Success && int.TryParse(match.Value, out int year))
            {
                return year;
            }

            // CC (年号なし) = 2013相当
            if (folderName.Contains("CC") && !folderName.Contains("20"))
            {
                return 2013;
            }

            // CS6 = 2012, CS5 = 2010, etc.
            if (folderName.Contains("CS6")) return 2012;
            if (folderName.Contains("CS5")) return 2010;
            if (folderName.Contains("CS4")) return 2009;
            if (folderName.Contains("CS3")) return 2007;

            return 0;
        }

        /// <summary>
        /// After Effectsを起動する
        /// </summary>
        private static void LaunchAfterEffects(string exePath, string projectPath, string debugVersionParams)
        {
             // Ctrlキー押下時はデバッグモード（AEを起動せずバージョン情報を表示）
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                ShowMessage($"AE version : {debugVersionParams}\r\r{projectPath}");
                return;
            }

            if (!File.Exists(exePath))
            {
                ShowMessage($"実行可能なAfter Effectsが見つかりません。\rPath: {exePath}\rProject: {projectPath}\rDetected Version: {debugVersionParams}");
                return;
            }

            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"\"{projectPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                ErrorDialog = true,
                ErrorDialogParentHandle = IntPtr.Zero,
                 // WindowStyle = ProcessWindowStyle.Normal // Default
            };

            Process.Start(psi);
        }

        private static void ShowMessage(string message)
        {
            MessageBox.Show(message, $"AEautoLauncher Version {Application.ProductVersion}");
        }
    }
}
