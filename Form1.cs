using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Deployment.Application;
using System.Web;
using System.IO;


namespace AEautoLauncher
{
    public partial class AEautoLauncher : Form
    {
        public AEautoLauncher()
        {
            InitializeComponent();
            Hikisuu_get();
        }

        public void Hikisuu_get()
        {
            string strProgramFilesX86Adobe = "C:\\Program Files (x86)\\Adobe\\";
            string strProgramFilesX64Adobe = "C:\\Program Files\\Adobe\\Adobe After Effects ";
            string strAfterEffectsLastPath = "\\Support Files\\AfterFX.exe";
            string strAEfullpath;
            string strAEversion;
            int aeversion;

            //コマンドラインを配列で取得する
            string[] cmds = System.Environment.GetCommandLineArgs();

            if (cmds.Length > 2)
            {
                MessageBox.Show(text: "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                    + "\r複数のファイル選択には対応していません");
            }
            else if (cmds.Length == 2)
            {
                FileStream rfs; // = null;
                rfs = new FileStream(cmds[1], FileMode.Open, FileAccess.Read, FileShare.Read);

                byte[] bytes = BinaryRead(rfs, 0x00000000, 0x00000030);

                //				MessageBox.Show(bytes[21].ToString());
                if(bytes[0x18]!=0x68)
                {
                    // CS5以前
                    aeversion = (((bytes[0x18] << 1) & 0xF8) + ((bytes[0x19] >> 3) & 0x07));
                    strAEversion = aeversion                                            // version
                        + "." + (((bytes[0x19] << 1) & 0x0E) + ((bytes[0x1A] >> 7)))    // minor
                        + "." + (((bytes[0x1A] >> 3) & 0x0F));                           // build
// 間違ってるかもしれない                       + "." + bytes[0x17];   // revision
                }
                else
                {
                    // CS6以降
                    aeversion = (((bytes[0x24] << 1) & 0xF8) + ((bytes[0x25] >> 3) & 0x07));
                    strAEversion = aeversion                                            // version
                        + "." + (((bytes[0x25] << 1) & 0x0E) + ((bytes[0x26] >> 7)))    // minor
                        + "." + (((bytes[0x26] >> 3) & 0x0F))                           // build
                        +"." + bytes[0x17];                                             // revision
                    if ((bytes[0x25] & 0x40) == 0)
                    {
                        strAEversion += "(Win)";
                    }
                    else
                    {
                        strAEversion += "(Mac)";
                    }
                }

                switch (aeversion) // macは0?000000
                {
                    case 5:
                        if (strAEversion.Substring(1, 3) != "5.5")
                        {
                            strAEfullpath = strProgramFilesX86Adobe + "After Effects 5.0" + strAfterEffectsLastPath;
                        }
                        else
                        {
                            strAEfullpath = strProgramFilesX86Adobe + "After Effects 5.5" + strAfterEffectsLastPath;
                        }
                        break;

                    case 6:
                        if (strAEversion.Substring(1, 3) != "6.5")
                        {
                            strAEfullpath = strProgramFilesX86Adobe + "After Effects 6.0" + strAfterEffectsLastPath;
                        }
                        else
                        {
                            strAEfullpath = strProgramFilesX86Adobe + "After Effects 6.5" + strAfterEffectsLastPath;
                        }
                        break;

                    case 7:
                        strAEfullpath = strProgramFilesX86Adobe + "After Effects 7.0" + strAfterEffectsLastPath;
                        break;

                    case 8:
                        strAEfullpath = strProgramFilesX86Adobe + "Adobe After Effects CS3" + strAfterEffectsLastPath;
                        strAEversion += " [CS3]";
                        break;

                    case 9:
                        strAEfullpath = strProgramFilesX86Adobe + "Adobe After Effects CS4" + strAfterEffectsLastPath;
                        strAEversion += " [CS4]";
                        break;

                    case 10:
                        if (strAEversion.Substring(1, 4) != "10.5")
                        {
                            strAEfullpath = strProgramFilesX64Adobe + "CS5" + strAfterEffectsLastPath;
                        }
                        else
                        {
                            strAEfullpath = strProgramFilesX64Adobe + "CS5.5" + strAfterEffectsLastPath;
                        }
                        strAEversion += " [CS5]";
                        break;

                    case 11:

                        strAEfullpath = strProgramFilesX64Adobe + "CS6" + strAfterEffectsLastPath;
                        strAEversion += " [CS6]";
                        break;

                    case 12:
                        strAEfullpath = strProgramFilesX64Adobe + "CC" + strAfterEffectsLastPath;
                        strAEversion += " [CC]";
                        break;

                    case 13: //13.8.1.38
                        strAEfullpath = strProgramFilesX64Adobe + "CC 2015.3" + strAfterEffectsLastPath;
                        strAEversion += " [CC 2015]";
                        break;


                    default:
                        strAEfullpath = "UnKnown";
                        if ((aeversion > 13) & (aeversion < 17)) // ver.14(CC 2017)-ver.16(CC 2019)まで
                        {
                            strAEfullpath = strProgramFilesX64Adobe + "CC " + (2003 + aeversion) + strAfterEffectsLastPath;
                        }
                        else if (aeversion > 16) // ver.17(CC 2020)から
                        {
                            strAEfullpath = strProgramFilesX64Adobe + (2003 + aeversion) + strAfterEffectsLastPath;
                        }
                        break;

                }

                if (strAEfullpath == "UnKnown")
                {
                    strAEfullpath = strProgramFilesX64Adobe + "2020" + strAfterEffectsLastPath;
                    AE_UnknownVersion(strAEfullpath, cmds[1], strAEversion);
                }
                else
                {
                    AE_exe(strAEfullpath, cmds[1], strAEversion);
                }

            }
            else
            {
                MessageBox.Show("AE5.0以降に対応\rフォルダはデフォルト決め打ち\r拡張子AEPの関連づけをAEautoLauncherにしてください。",
                    "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
            }
        }

        public static byte[] BinaryRead(FileStream binFileStream, long address, int length)
        {
            // byte型変数を格納できるListコレクションを宣言
            List<byte> data = new List<byte>();
            // ファイルストリームとアドレスのチェック
            if (binFileStream != null && address > -1)
            {
                // ファイルの最終アドレスより指定アドレスが大きい場合、
                // 空のバイト型配列を返す
                if (binFileStream.Length - 1 < address)
                {
                    return data.ToArray();
                }
                // 開始アドレス + 読み取り範囲の値がファイルの最終アドレスを
                // 超える場合、読み取り範囲をファイルの最終アドレス迄にする
                int readLength = length;
                if (address + readLength > binFileStream.Length)
                {
                    readLength = (int)(binFileStream.Length - address);
                }
                // バイナリ読み込み
                BinaryReader binReader = new BinaryReader(binFileStream);
                // 指定したアドレスに読み込み位置を移動
                binFileStream.Seek(address, SeekOrigin.Begin);
                // 読み込み
                data.AddRange(binReader.ReadBytes(readLength));
                // バイト型配列にして返す
                return data.ToArray();
            }
            else
            {
                // 空のバイト型配列を返す
                return data.ToArray();
            }
        }

        public void AE_exe(string strAEfullpath, string aep, string strAEversion)
        {
            // ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo hPsInfo = (
                new System.Diagnostics.ProcessStartInfo()
            );
            //			MessageBox.Show(ae_program);
            hPsInfo.FileName = strAEfullpath;

            // コントロールキーが押されている場合はヴァージョンチェック機能
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                MessageBox.Show("AE version : " + strAEversion + "\r\r" + @aep,
                    "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
            }

            // 実行ファイルがあるか？
            else if (File.Exists(hPsInfo.FileName))
            {


                // コマンドライン引数を設定する
                hPsInfo.Arguments = "\"" + @aep + "\"";

                // 新しいウィンドウを作成するかどうかを設定する (初期値 false)
                hPsInfo.CreateNoWindow = true;

                // シェルを使用するかどうか設定する (初期値 true)
                hPsInfo.UseShellExecute = false;

                // 起動できなかった時にエラーダイアログを表示するかどうかを設定する (初期値 false)
                hPsInfo.ErrorDialog = true;

                // エラーダイアログを表示するのに必要な親ハンドルを設定する
                hPsInfo.ErrorDialogParentHandle = this.Handle;

                // アプリケーションを起動する時の動詞を設定する
                hPsInfo.Verb = "Open";

                // 起動ディレクトリを設定する
                //			hPsInfo.WorkingDirectory = @"C:\Hoge\";

                // 起動時のウィンドウの状態を設定する
                hPsInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;     //通常
                                                                                        //			hPsInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;     //非表示
                                                                                        //			hPsInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized;  //最小化
                                                                                        //			hPsInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized;  //最大化

                // ProcessStartInfo を指定して起動する
                System.Diagnostics.Process.Start(hPsInfo);
            }
            else
            {
                MessageBox.Show("aepは実行ファイルの場所が違うので起動できません。\r" + hPsInfo.FileName + "\r" + @aep +"\rAE version: " + strAEversion,
                    "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());

            }
        }

        public void AE_UnknownVersion(string strAEfullpath, string aep, string strAEversion)
        {
            DialogResult result = MessageBox.Show("AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                + "\rバージョン不明ですがCC(2020)で起動ますか？\rAE version :" + strAEversion,
                "AEautoLauncher", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.OK)
            {
                AE_exe(strAEfullpath, aep, strAEversion);
            }
        }

        private void AEautoLauncher_Load(object sender, EventArgs e)
        {
            Close();
        }

    }
}
