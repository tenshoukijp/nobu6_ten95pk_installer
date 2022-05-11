using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Resources;
using System.Windows.Forms;

internal class Tenshouki95InstForm : Form
{
    private static string strCDDriveName;
    private static string strNoCDErrMsg = "天翔記 with PK の CDをドライブに挿入して下さい";

    private static string strSystemDriveName;  // ウィンドウズが入っている(多分HDDかSSDの)ドライブ名
    private static string strDstDir = "_Tenshouki95\\";

    private Label lbStatusMsg; // ステータス表示用

    private ProgressBar pbCopy;
    private static int iStatus = 0;

    private byte[] DDrawDLL;
    private byte[] N6PAudioDLL;
    private byte[] Tenshou95EXE;

    private int iKaoSwapFileSize = 1332;

    public Tenshouki95InstForm()
    {
        try
        {
            this.Width = 350;
            this.Height = 200;

            // プログレスバー
            ProgressBar progressBar = new ProgressBar();
            this.pbCopy = progressBar;
            progressBar.Left = 20;
            this.pbCopy.Top = 20;
            this.pbCopy.Width = 280;
            this.pbCopy.Height = 20;

            // ステータスの文字列
            Label label = new Label();
            this.lbStatusMsg = label;
            label.Left = 50;
            this.lbStatusMsg.Top = 50;
            this.lbStatusMsg.Width = 280;
            this.lbStatusMsg.Height = 80;

            Controls.Add(this.pbCopy);
            Controls.Add(this.lbStatusMsg);

            // このプロジェクトのアセンブリのタイプを取得。
            Assembly assembly = base.GetType().Assembly;
            ResourceManager r = new ResourceManager(string.Format("{0}.Ten95PK_installerRes", assembly.GetName().Name), assembly);

            this.Icon = (Icon)r.GetObject("icon");

            byte[] array1 = (byte[])r.GetObject("DDrawDLL");
            this.DDrawDLL = this.UnCompressBytes(array1);

            byte[] array2 = (byte[])r.GetObject("N6pAudioDLL");
            this.N6PAudioDLL = this.UnCompressBytes(array2);

            byte[] array3 = (byte[])r.GetObject("Tenshou95EXE");
            this.Tenshou95EXE = this.UnCompressBytes(array3);

            this.Shown += new EventHandler(Form_Shown);
            this.Closed += new EventHandler(Form_Closed);
        }
        catch
        {
            this.Dispose(true);
            throw;
        }
    }


    private Timer tmrClock;

    // 起動後１秒ごとにトライ。
    private void Form_Shown(object sender, EventArgs e)
    {
        // タイマーのインスタンス生成とイベントハンドラの追加
        this.tmrClock = new Timer();
        this.tmrClock.Tick += new EventHandler(this.tmrClock_Tick);

        // タイマーの間隔設定と始動
        this.tmrClock.Interval = 1000;
        this.tmrClock.Enabled = true;
    }

    // 実質的なメイン部分。ドライブをチェックし、存在したら、メイン処理部分を実行
    private void tmrClock_Tick(object sender, EventArgs e)
    {
        if (iStatus == 0)
        {
            this.tmrClock.Enabled = false;
            if (Tenshouki95ExistCheck())
            {
                this.tmrClock.Enabled = true;
                iStatus++;
            }
        }
        else
        {
            if (iStatus == 1)
            {
                iStatus++;
                this.Tenshouki95InstallExecute();
            }
        }
    }
    private void Form_Closed(object sender, EventArgs e)
    {
        // タイマーの停止
        this.tmrClock.Enabled = false;
    }

    private static bool CheckValidLogicalVolumeInfomation(DriveInfo d)
    {
        try {
            if (d.DriveType != DriveType.CDRom)
            {
                return false;
            }

            string volumeLabel = d.VolumeLabel;
            if (volumeLabel.Length >= 1 && volumeLabel.Contains("TENSHOUKI95"))
            {
                return true;
            }
        }
        catch (Exception)
        {
        }

        return false;
    }

    private string CheckValidLogicalDrive()
    {
        DriveInfo[] drives = DriveInfo.GetDrives();

        foreach (DriveInfo d in drives)
        {
            if (CheckValidLogicalVolumeInfomation(d))
            {
                return d.Name; // 天翔記ドライブがあった。"C:\"などの文字列かString::Empty
            }
        }
        return string.Empty; // 無かった
    }

    private bool Tenshouki95ExistCheck()
    {
        strCDDriveName = this.CheckValidLogicalDrive();
        if (strCDDriveName == string.Empty)
        {
            MessageBox.Show(strNoCDErrMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            this.Close();
            return false;
        }
        if (!File.Exists(strCDDriveName + "SETUP\\BFILE.N6P"))
        {
            MessageBox.Show(strNoCDErrMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            this.Close();
            return false;
        }
        if (!File.Exists(strCDDriveName + "SETUP\\SNDATA.N6P"))
        {
            MessageBox.Show(strNoCDErrMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            this.Close();
            return false;
        }

        strSystemDriveName = Environment.GetEnvironmentVariable("SystemDrive") + "\\";
        this.lbStatusMsg.Text = strCDDriveName + "ドライブに天翔記CDを発見しました！\n\n" +
                                strSystemDriveName + strDstDir + " フォルダに\nインストールを試みます。\n" +
                                                "インストール後、ご自身で\n" +
                                                "都合の良いフォルダへと移動して下さい。";
        return true;
    }

    private void Tenshouki95InstallExecute()
    {
        string[] strSrcFileNameList = Directory.GetFiles(strCDDriveName + "SETUP");
        string[] strDstFileNameList = Directory.GetFiles(strCDDriveName + "SETUP");
        for (int f = 0; f < strDstFileNameList.Length; f++)
        {
            strDstFileNameList[f] = strDstFileNameList[f].Replace(".95", ".EXE");
            strDstFileNameList[f] = strDstFileNameList[f].Replace(".EX_", ".EXE");
            strDstFileNameList[f] = strDstFileNameList[f].Replace(".N6_", ".N6P");
            strDstFileNameList[f] = strDstFileNameList[f].Replace("UNKOEI", "UNEOEI_");
            strDstFileNameList[f] = Path.GetFileName(strDstFileNameList[f]);
            strDstFileNameList[f] = strSystemDriveName + strDstDir + strDstFileNameList[f];
        }

        // プログレスバーの初期化
        this.pbCopy.Minimum = 0;
        this.pbCopy.Maximum = strSrcFileNameList.Length - 1;
        this.pbCopy.Value = 0;

        if (!Directory.Exists(strSystemDriveName + strDstDir))
        {
            Directory.CreateDirectory(strSystemDriveName + strDstDir);
        }

        string strErrorFile = "";
        for (int f = 0; f < strSrcFileNameList.Length; f++)
        {
            // コピー中のエラー対処
            try
            {
                string src = strSrcFileNameList[f];
                string dst = strDstFileNameList[f];

               
                // kaoswap.n6pなら、1332バイトで0で並んだファイルを書き出すだけ。
                if (dst.Contains("KAOSWAP.N"))
                {
                    byte[] dest = new byte[iKaoSwapFileSize];
                    File.WriteAllBytes(dst, dest);
                }
                else
                {
                    File.Copy(src, dst, true);
                }


                //読み取り専用属性を削除する
                FileAttributes attributes = File.GetAttributes(dst);
                File.SetAttributes(dst, attributes & ~FileAttributes.ReadOnly);

                //ProgressBar1の値を変更する
                this.pbCopy.Value = f;
            }
            catch (UnauthorizedAccessException)
            {
                strErrorFile = strDstFileNameList[f]; // エラーファイル控えておく
            }
        }

        // ここまでエラーがないなら、DDraw.dllを出力
        if (strErrorFile == "")
        {
            // ------------------------ DDraw.dll
            try
            {
                string fullDllName = strSystemDriveName + strDstDir + "DDraw.dll";

                // ファイルを作成して書き込む。ファイルが存在しているときは、上書きする
                FileStream fs = new FileStream(fullDllName, FileMode.Create, FileAccess.Write);

                fs.Write(DDrawDLL, 0, DDrawDLL.Length);
                fs.Close();


                DateTime dt = new DateTime(2010, 3, 31, 6, 0, 0);

                //作成日時の設定
                File.SetCreationTime(fullDllName, dt);

                //更新日時の設定
                File.SetLastWriteTime(fullDllName, dt);
            }
            catch (UnauthorizedAccessException)
            {
                strErrorFile = "DDraw.dll";
            }

            // ------------------------ N6PAudio.dll
            try
            {
                string fullDllName = strSystemDriveName + strDstDir + "N6PAudio.dll";

                // ファイルを作成して書き込む。ファイルが存在しているときは、上書きする
                FileStream fs = new FileStream(fullDllName, FileMode.Create, FileAccess.Write);
                fs.Write(N6PAudioDLL, 0, N6PAudioDLL.Length);
                fs.Close();

                DateTime dt = new DateTime(2002, 7, 26, 16, 30, 44);
                File.SetCreationTime(fullDllName, dt);
                File.SetLastWriteTime(fullDllName, dt);
            }
            catch (UnauthorizedAccessException)
            {
                strErrorFile = "N6PAudio.dll";
            }


            // ------------------------ Tenshou.exe
            try
            {
                string fullExeName = strSystemDriveName + strDstDir + "Tenshou.exe";

                // ファイルを作成して書き込む。ファイルが存在しているときは、上書きする
                FileStream fs = new FileStream(fullExeName, FileMode.Create, FileAccess.Write);
                fs.Write(Tenshou95EXE, 0, Tenshou95EXE.Length);
                fs.Close();


                DateTime dt = new DateTime(2002, 11, 19, 18, 3, 40);
                File.SetCreationTime(fullExeName, dt);
                File.SetLastWriteTime(fullExeName, dt);
            }
            catch (UnauthorizedAccessException)
            {
                strErrorFile = "Tenshou.exe";
            }
        }
        if (strErrorFile != "")
        {
            MessageBox.Show("コピー中でエラーが発生しました。\n" + "ファイル名:" + strErrorFile, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }
        else
        {
            MessageBox.Show("インストールが完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }
    }

    private byte[] UnCompressBytes(byte[] byteComp)
    {
        byte[] buf = new byte[1024]; // 1Kbytesずつ処理する

        // 入力ストリーム
        MemoryStream inStream = new MemoryStream(byteComp);
        MemoryStream outStream = new MemoryStream();

        GZipStream decompStream = new GZipStream(inStream, CompressionMode.Decompress);

        while (true)
        {
            int num = decompStream.Read(buf, 0, buf.Length);
            if (num > 0)
            {
                outStream.Write(buf, 0, num);
            }
            else
            {
                break;
            }

        }

        return outStream.ToArray();
    }
}
