using ComponentFactory.Krypton.Toolkit;
using Microsoft.WindowsAPICodePack.Dialogs;
using Syncfusion.Windows.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace RTEditor
{
    public partial class InsertGraphDialog : KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);
        public InsertGraphDialog()
        {
            InitializeComponent();

            //タイトルバーのダークモード
            if (Properties.Settings.Default.IsUseDarkModeTitleBar == true)
            {
                const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
                DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, new[] { 1 }, sizeof(int));
            }
            else if (Properties.Settings.Default.IsUseDarkModeTitleBar == false)
            {
                const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
                DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, new[] { 0 }, sizeof(int));
            }
        }

        private void SeriesFromCSV(Chart chart)
        {
            // シリーズの名前を用意する
            string srsStr = "項目";
            // シリーズ用のインスタンスを生成する
            Series series = new Series(srsStr);
            // シリーズのポイント幅を設定する
            series["PointWidth"] = "0.5";
            // シリーズをチャートに追加する
            chart.Series.Add(series);
            // シリーズのポイントをクリアする
            chart.Series[srsStr].Points.Clear();
            // シリーズのチャートタイプを設定する
            chart.Series[srsStr].ChartType = SeriesChartType.Column;
            // CSVファイルを読み込むStreamReaderのインスタンスを生成する
            using (var sr = new StreamReader(CSVFilePath))
            {
                // CSVファイルの1行目を読み込む
                // ⇒ヘッダ行の1行をスキップするため、空で1行目を読み込む
                sr.ReadLine();
                // ファイルの終わりまで読み込むWhileループ
                while (!sr.EndOfStream)
                {
                    // 1つ行を取得して変数lineに格納する
                    var line = sr.ReadLine();
                    // 読み込んだ行をカンマで分割して配列valuesに格納する
                    var values = line.Split(',');
                    // 2列目(名前)を取得して変数nameに格納する
                    string name = values[1].Trim('"');
                    // 4列目(在庫)を取得して変数stockに格納する
                    int stock = int.Parse(values[3].Trim('"'));
                    // シリーズにデータポイントを追加する
                    chart.Series[srsStr].Points.AddXY(name, stock);
                }
            }
        }
        string CSVFilePath;
        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "CSVファイル (*.csv)|*.csv", Multiselect = false })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    chart.Series.Clear();
                    CSVFilePath = openFileDialog.FileName;
                    SeriesFromCSV(chart);
                    this.Text = "グラフの挿入 - " + openFileDialog.FileName;
                }
            }
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            CaptureControl(chart);
        }
        public static Bitmap CaptureControl(Control ctrl)
        {
            Bitmap bmp = new Bitmap(ctrl.Width, ctrl.Height);
            ctrl.DrawToBitmap(bmp, new Rectangle(0, 0, ctrl.Width, ctrl.Height));
            Clipboard.SetImage(bmp);
            return bmp;
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            MessageBoxAdv.Show("CSVファイルの読み込み方法について\n\n1: CSVファイルを選択します。\nCSVファイルの内容はカンマ区切りで項目を区切り下記のように指定しExcelなどでCSVファイル(UTF-8 BOM付き)として保存します。\n例:\r\n\"番号\",\"名称\",\"単価\",\"合計売上\",\"在庫\"\r\n\"1\",\"リンゴ\",\"100\",\"100000\",\"156\"\r\n\"2\",\"梨\",\"100\",\"200000\",\"100\"\r\n\"3\",\"イチゴ\",\"300\",\"300000\",\"0\"\r\n\n\n2: グラフの種類を選択します。\n3: [挿入]ボタンをクリックします。\n\n※グラフが画像としてリッチテキストエディタに挿入されます。貼り付け後のグラフのスタイルは変更できませんのでご注意ください。", "CSVファイルの読み込みについて", MessageBoxButtons.OK);
        }

        private void kryptonComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            chart.Palette = (ChartColorPalette)kryptonComboBox1.SelectedIndex;
        }

        private void kryptonComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            chart.Series[0].ChartType = (SeriesChartType)kryptonComboBox2.SelectedIndex;
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void InsertGraphDialog_Load(object sender, EventArgs e)
        {
            //テーマ適用

            //既定のテーマ
            if (Properties.Settings.Default.Theme == "Global")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Global;
            }
            //Professional-Systemテーマ
            else if (Properties.Settings.Default.Theme == "ProfessionalSystem")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.ProfessionalSystem;
            }
            //Professional-Office2003テーマ
            else if (Properties.Settings.Default.Theme == "ProfessionalOffice2003")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.ProfessionalOffice2003;
            }
            //Office2007Blueテーマ
            else if (Properties.Settings.Default.Theme == "Office2007Blue")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007Silverテーマ
            else if (Properties.Settings.Default.Theme == "Office2007Silver")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007Blackテーマ
            else if (Properties.Settings.Default.Theme == "Office2007Black")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010Blueテーマ
            else if (Properties.Settings.Default.Theme == "Office2010Blue")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010Silverテーマ
            else if (Properties.Settings.Default.Theme == "Office2010Silver")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010Blackテーマ
            else if (Properties.Settings.Default.Theme == "Office2010Black")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            //SparkleBlueテーマ
            else if (Properties.Settings.Default.Theme == "SparkleBlue")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.SparkleBlue;
            }
            //SparkleOrangeテーマ
            else if (Properties.Settings.Default.Theme == "SparkleOrange")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.SparkleOrange;
            }
            //SparklePurpleテーマ
            else if (Properties.Settings.Default.Theme == "SparklePurple")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.SparklePurple;
            }

            if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == true)
            {
                kryptonPalette1.AllowFormChrome = ComponentFactory.Krypton.Toolkit.InheritBool.False;
            }
            else if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == false)
            {
                kryptonPalette1.AllowFormChrome = ComponentFactory.Krypton.Toolkit.InheritBool.True;
            }
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
        }
    }
}
