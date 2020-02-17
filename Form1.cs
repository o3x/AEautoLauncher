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
            string strAfterEffectsLastPass = "\\Support Files\\AfterFX.exe";
            string strAEfullpath = "";

            //コマンドラインを配列で取得する
            string[] cmds = System.Environment.GetCommandLineArgs();

            if (cmds.Length > 2)
            {
                MessageBox.Show(text: "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                    + "\r複数のファイル選択には対応していません");
            }
            else if (cmds.Length > 1)
            {
                FileStream rfs; // = null;
                rfs = new FileStream(cmds[1], FileMode.Open, FileAccess.Read, FileShare.Read);

                byte[] bytes = BinaryRead(rfs, 0x00000000, 0x00000030);

                //				MessageBox.Show(bytes[21].ToString());

                switch (bytes[0x15])
                {
                    case 0x44:
                        strAEfullpath = strProgramFilesX86Adobe + "After Effects 6.5" + strAfterEffectsLastPass;
                        break;

                    case 0x49:
                        strAEfullpath = strProgramFilesX86Adobe + "Adobe After Effects CS3" + strAfterEffectsLastPass;
                        break;

                    case 0x4A:
                        strAEfullpath = strProgramFilesX86Adobe + "Adobe After Effects CS4" + strAfterEffectsLastPass;
                        break;

                    case 0x4C:
                        strAEfullpath = strProgramFilesX64Adobe + "CS5" + strAfterEffectsLastPass;
                        break;

                    default:

                        switch (bytes[0x25] % 0x40) // macは0?000000
                        {

                            case 0x12:

                                switch (bytes[0x21])
                                {
                                    case 0x4C:
                                        strAEfullpath = strProgramFilesX64Adobe + "CS5" + strAfterEffectsLastPass;
                                        break;

                                    case 0x4E:
                                        strAEfullpath = strProgramFilesX64Adobe + "CS5.5" + strAfterEffectsLastPass;
                                        break;
                                }
                                break;

                            case 0x18:

                                switch (bytes[0x21])
                                {
                                    case 0x4D:
                                    case 0x4E:
                                        strAEfullpath = strProgramFilesX64Adobe + "CS5.5" + strAfterEffectsLastPass;
                                        break;

                                    case 0x51:
                                        strAEfullpath = strProgramFilesX64Adobe + "CS6" + strAfterEffectsLastPass;
                                        break;
                                }
                                break;

                            case 0x20:
                            case 0x21:
                                strAEfullpath = strProgramFilesX64Adobe + "CC" + strAfterEffectsLastPass;
                                break;

                            case 0x28:
                            case 0x29:
                                strAEfullpath = strProgramFilesX64Adobe + "CC" + strAfterEffectsLastPass;
                                break;

                            case 0x2A:
                            case 0x2B:
                            case 0x2C: //13.8.1.38
                                strAEfullpath = strProgramFilesX64Adobe + "CC 2015.3" + strAfterEffectsLastPass;
                                break;

                            case 0x31: //V.14
                                strAEfullpath = strProgramFilesX64Adobe + "CC 2017" + strAfterEffectsLastPass;
                                break;

                            case 0x38: //V.15
                                strAEfullpath = strProgramFilesX64Adobe + "CC 2018" + strAfterEffectsLastPass;
                                break;

                            case 0x00: //cc2019
                                strAEfullpath = strProgramFilesX64Adobe + "CC 2019" + strAfterEffectsLastPass;
                                break;

                            case 0x08: //cc2020
                                strAEfullpath = strProgramFilesX64Adobe + "2020" + strAfterEffectsLastPass;
                                break;

                            default:
                                strAEfullpath = "UnKnown";
                                break;

                        }
                        break;
                }
                if (strAEfullpath == "UnKnown")
                {
                    strAEfullpath = strProgramFilesX64Adobe + "CC 2018" + strAfterEffectsLastPass;
                    AE_UnknownVersion(strAEfullpath, cmds[1], bytes);
                }
                else
                {
                    AE_exe(strAEfullpath, cmds[1]);
                }

                /*				string filename = System.IO.Path.GetFileName(cmds[1]);
                                if (filename.IndexOf("comp") > 0)
                                {
                                    AE_exe("cs3", cmds[1]);
                                }
                                else
                                {
                                    //AE CS4実行
                                    AE_exe("cs4", cmds[1]);
                                }
                 */

            }
            else
            {
                MessageBox.Show(text: "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                    + "\r暫定仕様です。\rAE6.5～CC 2020\rフォルダはデフォルト決め打ち\r拡張子AEPの関連づけをAEautoLauncherにしてください。");
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

        public void AE_exe(string strAEfullpath, string aep)
        {
            // ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo hPsInfo = (
                new System.Diagnostics.ProcessStartInfo()
            );
            //			MessageBox.Show(ae_program);
            hPsInfo.FileName = strAEfullpath;

            // 実行ファイルがあるか？
            if (File.Exists(hPsInfo.FileName))
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
                if ((Control.ModifierKeys & Keys.Control) != Keys.Control)
                {
                    System.Diagnostics.Process.Start(hPsInfo);

                }
                else
                {
                    MessageBox.Show("AfterEffects\r" + hPsInfo.FileName + "\r\rAEPファイル\r" + @aep, "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
                }
            }
            else
            {
                MessageBox.Show(text: "AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                    + "\raepは実行ファイルの場所が違うので起動できません。\r"
                    + hPsInfo.FileName + "\r" + @aep);

            }
        }

        public void AE_UnknownVersion(string strAEfullpath, string aep, byte[] bytes)
        {
            DialogResult result = MessageBox.Show("AEautoLauncher Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                + "\rバージョン不明ですがCC(2018)で起動してみます。\r確認コード:" + bytes[0x15].ToString("x2") + ":" + bytes[0x21].ToString("x2"),
                "AEautoLauncher", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.OK)
            {
                AE_exe(strAEfullpath, aep);
            }
        }

        private void AEautoLauncher_Load(object sender, EventArgs e)
        {
            Close();
        }

    }
}
