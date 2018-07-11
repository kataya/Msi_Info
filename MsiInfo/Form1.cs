using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsInstaller;

namespace MsiInfo
{

    /// <summary>
    /// インストールやアンインストールのときに必要な MSI や MSPの情報を[orca]なしで確認するためのサンプル
    /// ・PowerShellでの実装
    ///      https://marcinotorowski.com/2018/03/04/enumerating-installed-msi-products-with-powershell/
    /// ・pinvoke:: DLL importでの実装
    ///      https://www.neowin.net/forum/topic/968772-vbnet-msi-manipulation/
    /// </summary>

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog fld = new OpenFileDialog();
            fld.InitialDirectory = Environment.CurrentDirectory;
            fld.Filter = "MSI ファイル (*.msi)|*.msi|MSP ファイル (*.msp)|*.msp";

            if (fld.ShowDialog() == DialogResult.OK) {
                string selfile = fld.FileName;
                string ext = System.IO.Path.GetExtension(selfile);
                string info = string.Empty;
                if (ext.Equals(".msi")){
                    //ProductName, ProductCode, UpgradeCode, Manufacturer, ARPHELPLINK, ARPCOMMENTS, ARPCONTACT, ARPURLINFOABOUT and ARPURLUDATEINFO
                    info = GetMsiProperty(selfile, "ProductCode");
                }
                else if (ext.Equals(".msp")) {
                    info = GetMspSummaryInfo(selfile);
                }

                this.textBox1.Text = info;
                //MessageBox.Show(this, info);
            }


        }

        /// <summary>
        /// MSP ファイルからアンインストール時に必要な情報を取得
        /// </summary>
        /// <param name="mspFile"></param>
        /// <returns></returns>
        public static string GetMspSummaryInfo(string mspFile)
        {
            string retVal = string.Empty;

            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Object installerObj = Activator.CreateInstance(classType);
            Installer installer = installerObj as Installer;

            // 読み込みのために msp ファイルをオープン
            // 0 - Read, 1 - Read/Write
            Database database = installer.OpenDatabase(mspFile, 32);

            int propCnt = database.SummaryInformation.PropertyCount;

            // orca での 「Patch Summary Information」画面 
            //retVal = string.Format("Targets: {0}", database.SummaryInformation.Property[7]); // [7]は Targets:

            //retVal += " , " + string.Format("Patch Code: {0}", database.SummaryInformation.Property[9]); // [9]は Patch Code:

            // アンインストール用のコマンドの形式で返す
            // msiexec / I < ProductGUID > MSIPATCHREMOVE =< PatchGUID >
            retVal = string.Format("msiexec /I '{0}' MSIPATCHREMOVE='{1}'", database.SummaryInformation.Property[7], database.SummaryInformation.Property[9]);
            retVal = retVal.Replace("'", "\"");

            return retVal;
        }


        /// <summary>
        /// Msi ファイルからアンインストール時に必要な情報を取得
        /// 参考1: Read Properties from an MSI File
        ///   http://www.alteridem.net/2008/05/20/read-properties-from-an-msi-file/
        /// 
        /// 参考2: COM Microsoft Windows Installer Object Library を利用とあるが参照しようとするとエラー
        ///   %WINDIR%\system32\msi.dll をかわりに追加するといい
        ///   http://www.reinholdt.me/2015/07/29/problem-adding-a-c-reference-to-windowsinstaller-com-object/
        /// </summary>
        /// <param name="msiFile"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        public static string GetMsiProperty(string msiFile, string property)
        {
            string retVal = string.Empty;


            // インストーラのインスタンスを作成
            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Object installerObj = Activator.CreateInstance(classType);
            Installer installer = installerObj as Installer;

            // 読み込みのために msi ファイルをオープン
            // 0 - Read, 1 - Read/Write
            Database database = installer.OpenDatabase(msiFile,0);

            // SQL文でプロパティをクエリして取得
            string sql = String.Format(
                "SELECT Value FROM Property WHERE Property ='{0}'", property);
            WindowsInstaller.View view = database.OpenView(sql);
            view.Execute(null);

            // 取得したレコードの読み込み
            Record record = view.Fetch();
            if (record != null)
                //retVal = string.Format("{0}: {1}", property, record.get_StringData(1));
                // アンインストール用のコマンドの形式で返す
                // mxiexec /X <PruductGUID>
                retVal = string.Format("mxiexec /X '{0}'", record.get_StringData(1));
                retVal = retVal.Replace("'", "\"");

            return retVal;
        }


    }
}
