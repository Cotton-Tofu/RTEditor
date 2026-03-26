using ComponentFactory.Krypton.Navigator;
using ComponentFactory.Krypton.Toolkit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RTEditor
{
    public partial class AIPromptAnswarDialog : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);

        public AIPromptAnswarDialog()
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

        

        //AIPromptAnswarDialog内のコントロールを外部から操作するためのプロパティ
        //AIPromptAnswarDialogのインスタンスを取得できるようにするためのプロパティ
        private static AIPromptAnswarDialog _AIPromptAnswarDialogInstance { get; set; }

        public static AIPromptAnswarDialog Instance
        {
            get { return _AIPromptAnswarDialogInstance; }
            set { _AIPromptAnswarDialogInstance = value; }
        }

        private void AIPromptAnswarDialog_Load(object sender, EventArgs e)
        {
            //AIPromptAnswarDialogのインスタンスを保存する
            _AIPromptAnswarDialogInstance = this;



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

            kryptonCheckBox1_CheckedChanged(sender, e);
        }

        private void AIPromptAnswarDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            //AIPromptAnswarDialogのインスタンスを破棄する
            Instance = null;
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(kryptonCheckBox1.Checked == true)
            {
                this.TopMost = true;
            }
            else if(kryptonCheckBox1.Checked == false)
            {
                this.TopMost = false;
            }
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            KryptonPage page = kryptonNavigator1.SelectedPage as KryptonPage;
            KryptonRichTextBox rtb = page.Controls.OfType<KryptonRichTextBox>().FirstOrDefault();
            rtb.Focus();
            rtb.Copy();
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            KryptonPage page = kryptonNavigator1.SelectedPage as KryptonPage;
            KryptonRichTextBox rtb = page.Controls.OfType<KryptonRichTextBox>().FirstOrDefault();
            rtb.Focus();
            rtb.SelectAll();
            rtb.ScrollToCaret();
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            KryptonPage page = kryptonNavigator1.SelectedPage as KryptonPage;
            KryptonRichTextBox rtb = page.Controls.OfType<KryptonRichTextBox>().FirstOrDefault();

            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb1 = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb1 == null) return;

            if (rtb.SelectedText == string.Empty)
            {
                //すべて貼り付け
                rtb1.AppendText(rtb.Text);
            }
            else
            {
                //選択箇所のみ貼り付け
                rtb1.AppendText(rtb.SelectedText);
            }
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            KryptonPage page = kryptonNavigator1.SelectedPage as KryptonPage;
            KryptonRichTextBox rtb = page.Controls.OfType<KryptonRichTextBox>().FirstOrDefault();
            using (SaveFileDialog saveFileDialog = new SaveFileDialog { Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt", Title = "保存する場所を選択..." })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    rtb.SaveFile(saveFileDialog.FileName);
            }
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox2.Checked == true)
            {
                this.Opacity = 0.8D;
            }
            else
            {
                this.Opacity = 100;
            }
        }

        private void kryptonContextMenuItem3_Click(object sender, EventArgs e)
        {
            KryptonPage page = kryptonNavigator1.SelectedPage as KryptonPage;
            KryptonRichTextBox rtb = page.Controls.OfType<KryptonRichTextBox>().FirstOrDefault();
            rtb.Focus();
            rtb.SelectAll();
        }
    }
}
