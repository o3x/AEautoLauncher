using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

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
                    // Fallback or Unknown handling
                    string defaultPath = ProgramFilesX64Adobe + @"Adobe After Effects 2020" + AfterEffectsExePath;
                    
                    // If the specific version path isn't found, try a known fallback or ask user
                    DialogResult result = MessageBox.Show(
                        $"バージョン不明または未インストールのバージョンです。\rCC(2020)で起動しますか？\r検出されたバージョン: {strVersionInfo}",
                        $"AEautoLauncher Version {Application.ProductVersion}",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.OK)
                    {
                        LaunchAfterEffects(defaultPath, aepPath, strVersionInfo);
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
                    byte[] bytes = br.ReadBytes(48); // Read header
                    if (bytes.Length < 48) return 0;

                    // Magic Number Check: RIFF ... Egg!
                    // RIFF = 0x52, 0x49, 0x46, 0x46 (0-3)
                    // RIFX = 0x52, 0x49, 0x46, 0x58 (0-3) - Big Endian variant
                    // Egg! = 0x45, 0x67, 0x67, 0x21 (8-11)
                    bool isRiff = (bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46);
                    bool isRifx = (bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x58);
                    bool isEgg  = (bytes[8] == 0x45 && bytes[9] == 0x67 && bytes[10] == 0x67 && bytes[11] == 0x21);

                    if ((!isRiff && !isRifx) || !isEgg)
                    {
                        return 0; // Not a valid AEP file
                    }

                    bool isCs6OrLater = (bytes[0x18] == 0x68);

                    if (!isCs6OrLater)
                    {
                        // CS5 and earlier
                        version = ((bytes[0x18] << 1) & 0xF8) + ((bytes[0x19] >> 3) & 0x07);
                        int minor = ((bytes[0x19] << 1) & 0x0E) + (bytes[0x1A] >> 7);
                        int build = (bytes[0x1A] >> 3) & 0x0F;
                        versionString = $"{version}.{minor}.{build}";
                    }
                    else
                    {
                        // CS6 and later
                        version = ((bytes[0x24] << 1) & 0xF8) + ((bytes[0x25] >> 3) & 0x07);
                        int minor = ((bytes[0x25] << 1) & 0x0E) + (bytes[0x26] >> 7);
                        int build = (bytes[0x26] >> 3) & 0x0F;
                        int revision = bytes[0x27];
                        versionString = $"{version}.{minor}.{build}.{revision}";

                        // Extract host version for additional info
                        int hostVer = ((bytes[0x14] << 1) & 0xF8) + ((bytes[0x15] >> 3) & 0x07);
                        int hostMinor = ((bytes[0x15] << 1) & 0x0E) + (bytes[0x16] >> 7);
                        int hostBuild = (bytes[0x16] >> 3) & 0x0F;
                        string hostVerString = $"{hostVer}.{hostMinor}.{hostBuild}.{bytes[0x17]}";

                        string platform = (bytes[0x25] & 0x40) == 0 ? "(Win)" : "(Mac)";
                        versionString += platform;

                        if (versionString != hostVerString)
                        {
                            versionString += $" [HostVersion:{hostVerString}]";
                        }
                    }
                }
            }
            catch
            {
                // Ignore read errors, return 0
            }

            return version;
        }

        private static string ResolveAePath(int version)
        {
            // Simple mapping for older versions
            if (version == 5) return ProgramFilesX86Adobe + @"After Effects 5.5" + AfterEffectsExePath; // Handling 5.0/5.5 logic simplified for now as per old logic 'if < 5.5' check was there, but assume 5.5 for simplicity or fallback
            if (version == 6) return ProgramFilesX86Adobe + @"After Effects 6.5" + AfterEffectsExePath;
            if (version == 7) return ProgramFilesX86Adobe + @"After Effects 7.0" + AfterEffectsExePath;
            if (version == 8) return ProgramFilesX86Adobe + @"Adobe After Effects CS3" + AfterEffectsExePath;
            if (version == 9) return ProgramFilesX86Adobe + @"Adobe After Effects CS4" + AfterEffectsExePath;
            
            // CS5 - CS6
            if (version == 10) return ProgramFilesX64Adobe + @"Adobe After Effects CS5.5" + AfterEffectsExePath; // Warning: Old logic had CS5 exception
            if (version == 11) return ProgramFilesX64Adobe + @"Adobe After Effects CS6" + AfterEffectsExePath;
            
            // CC versions
            if (version == 12) return ProgramFilesX64Adobe + @"Adobe After Effects CC" + AfterEffectsExePath;
            if (version == 13) return ProgramFilesX64Adobe + @"Adobe After Effects CC 2015.3" + AfterEffectsExePath; // Old logic split 2014/2015

            // Automatic mapping for CC 2017+ (v14+)
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

        private static void LaunchAfterEffects(string exePath, string projectPath, string debugVersionParams)
        {
             // Check for Control key for debug mode
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
