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
    public partial class InsertMathematicalSymbolsDialog : KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);
        public InsertMathematicalSymbolsDialog()
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

        private void kryptonPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void InsertMathematicalSymbolsDialog_Load(object sender, EventArgs e)
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

        private void InsertMathematicalSymbolsDialog_Activated(object sender, EventArgs e)
        {
        }

        private void InsertMathematicalSymbolsDialog_Deactivate(object sender, EventArgs e)
        {
        }

        public void InsertSymbolOrText(string symbolText)
        {
            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb == null) return;

            rtb.SelectionFont = new Font("Cambria Math", rtb.SelectionFont.Size);
            rtb.AppendText(symbolText);

        }


        //数字・算術・かっこ・単位・文字タブ

        //数字
        //1
        private void kryptonButton10_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton10.Text);
        }

        //2
        private void kryptonButton11_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton11.Text);
        }

        //3
        private void kryptonButton12_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton12.Text);
        }

        //4
        private void kryptonButton13_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton13.Text);
        }

        //5
        private void kryptonButton14_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton14.Text);
        }

        //6
        private void kryptonButton15_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton15.Text);
        }

        //7
        private void kryptonButton16_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton16.Text);
        }

        //8
        private void kryptonButton17_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton17.Text);
        }

        //9
        private void kryptonButton18_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton18.Text);
        }

        //0
        private void kryptonButton19_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton19.Text);
        }

        //.
        private void kryptonButton38_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton38.Text);
        }


        //算術
        //＋
        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton1.Text);
        }

        //−
        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton2.Text);
        }

        //÷
        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton3.Text);
        }

        //×
        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton4.Text);
        }

        //%
        private void kryptonButton37_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton37.Text);
        }

        // /(除法)
        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton5.Text);
        }

        //*
        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton6.Text);
        }

        //±
        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton7.Text);
        }

        //∓
        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton8.Text);
        }

        //=
        private void kryptonButton9_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton9.Text);
        }

        //√
        private void kryptonButton27_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton27.Text);
        }

        //∞
        private void kryptonButton28_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton28.Text);
        }

        //≈
        private void kryptonButton29_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton29.Text);
        }

        //≤
        private void kryptonButton30_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton30.Text);
        }

        //≥
        private void kryptonButton31_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton31.Text);
        }

        //<
        private void kryptonButton45_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton45.Text);
        }

        //>
        private void kryptonButton46_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton46.Text);
        }

        //≠
        private void kryptonButton47_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton47.Text);
        }


        //かっこ
        //{
        private void kryptonButton21_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton21.Text);
        }

        //}
        private void kryptonButton22_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton22.Text);
        }

        //(
        private void kryptonButton23_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton23.Text);
        }

        //)
        private void kryptonButton24_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton24.Text);
        }

        //[
        private void kryptonButton25_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton25.Text);
        }

        //]
        private void kryptonButton26_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton26.Text);
        }

        //単位
        //℃
        private void kryptonButton57_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton57.Text);
        }

        //℉
        private void kryptonButton58_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton58.Text);
        }

        //μ
        private void kryptonButton59_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton59.Text);
        }

        //n
        private void kryptonButton60_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton60.Text);
        }

        //m
        private void kryptonButton61_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton61.Text);
        }

        //f
        private void kryptonButton65_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton65.Text);
        }

        //k
        private void kryptonButton62_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton62.Text);
        }

        //g
        private void kryptonButton63_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton63.Text);
        }

        //t
        private void kryptonButton64_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton64.Text);
        }

        //文字
        //a
        private void kryptonButton142_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton142.Text);
        }

        //b
        private void kryptonButton143_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton143.Text);
        }

        //c
        private void kryptonButton144_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton144.Text);
        }

        //d
        private void kryptonButton145_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton145.Text);
        }

        //x
        private void kryptonButton146_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton146.Text);
        }

        //y
        private void kryptonButton147_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton147.Text);
        }

        //z
        private void kryptonButton148_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton148.Text);
        }

        //集合タブ

        //集合
        //∈
        private void kryptonButton77_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton77.Text);
        }

        //∋
        private void kryptonButton78_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton78.Text);
        }

        //⊆
        private void kryptonButton79_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton79.Text);
        }

        //⊇
        private void kryptonButton80_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton80.Text);
        }

        //⊂
        private void kryptonButton81_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton81.Text);
        }

        //⊃
        private void kryptonButton82_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton82.Text);
        }

        //⊄
        private void kryptonButton83_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton83.Text);
        }

        //⊅
        private void kryptonButton84_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton84.Text);
        }

        //⊊
        private void kryptonButton85_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton85.Text);
        }

        //⊋
        private void kryptonButton86_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton86.Text);
        }

        //∉
        private void kryptonButton87_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton87.Text);
        }


        //∌
        private void kryptonButton88_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton88.Text);
        }

        //論理タブ

        //論理
        //⇒
        private void kryptonButton89_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton89.Text);
        }

        //→
        private void kryptonButton90_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton90.Text);
        }

        //⊃
        private void kryptonButton91_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton91.Text);
        }

        //⇔
        private void kryptonButton92_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton92.Text);
        }

        //≡
        private void kryptonButton93_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton93.Text);
        }

        //↔
        private void kryptonButton98_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton94.Text);
        }

        //¬
        private void kryptonButton97_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton97.Text);
        }

        //˜
        private void kryptonButton96_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton96.Text);
        }

        //!
        private void kryptonButton95_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton95.Text);
        }

        //∧
        private void kryptonButton94_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton94.Text);
        }

        //·
        private void kryptonButton103_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton103.Text);
        }

        //⋅
        private void kryptonButton102_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton102.Text);
        }

        //&
        private void kryptonButton101_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton101.Text);
        }

        //∨
        private void kryptonButton100_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton100.Text);
        }

        //+
        private void kryptonButton99_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton99.Text);
        }

        //∥
        private void kryptonButton108_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton108.Text);
        }

        //⊕
        private void kryptonButton107_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton107.Text);
        }

        //⊻
        private void kryptonButton106_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton106.Text);
        }

        //⊤
        private void kryptonButton105_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton105.Text);
        }

        //T
        private void kryptonButton104_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton104.Text);
        }

        //1
        private void kryptonButton113_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton103.Text);
        }

        //⊥
        private void kryptonButton112_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton112.Text);
        }

        //F
        private void kryptonButton111_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton111.Text);
        }

        //0
        private void kryptonButton110_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton110.Text);
        }

        //∀
        private void kryptonButton109_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton109.Text);
        }

        //∃
        private void kryptonButton118_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton108.Text);
        }

        //∃!
        private void kryptonButton117_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton117.Text);
        }

        //≡
        private void kryptonButton116_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton116.Text);
        }

        //:⇔
        private void kryptonButton115_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton115.Text);
        }

        //( )
        private void kryptonButton114_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton114.Text);
        }

        //⊢
        private void kryptonButton35_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton35.Text);
        }

        //⊨
        private void kryptonButton119_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton119.Text);
        }

        //微積タブ

        //微分
        //′
        private void kryptonButton36_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton36.Text);
        }

        //″
        private void kryptonButton120_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton120.Text);
        }

        //∂
        private void kryptonButton121_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton121.Text);
        }


        //d
        private void kryptonButton122_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton122.Text);
        }

        //ẋ
        private void kryptonButton123_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton123.Text);
        }

        //ÿ
        private void kryptonButton124_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton124.Text);
        }

        //積分
        //∫
        private void kryptonButton125_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton125.Text);
        }

        //∬
        private void kryptonButton126_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton126.Text);
        }

        //∭
        private void kryptonButton127_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton127.Text);
        }

        //∮
        private void kryptonButton129_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton129.Text);
        }

        //∯
        private void kryptonButton130_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton130.Text);
        }

        //∰
        private void kryptonButton131_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton131.Text);
        }

        //極限
        //lim
        private void kryptonButton132_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton132.Text);
        }

        //→
        private void kryptonButton133_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton133.Text);
        }

        //∞
        private void kryptonButton134_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton134.Text);
        }

        //演算子・関数
        //∇
        private void kryptonButton137_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton137.Text);
        }

        //∆
        private void kryptonButton136_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton136.Text);
        }

        //√
        private void kryptonButton135_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton135.Text);
        }

        //∑
        private void kryptonButton138_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton138.Text);
        }

        //その他の数学記号タブ

        //その他
        //Π
        private void kryptonButton139_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton139.Text);
        }

        //π
        private void kryptonButton32_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton32.Text);
        }

        //Ω
        private void kryptonButton140_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton140.Text);
        }

        //ω
        private void kryptonButton141_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton141.Text);
        }

        //プログラム用記号

        //プログラム記号
        //==
        private void kryptonButton39_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton39.Text);
        }

        //!-
        private void kryptonButton40_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton40.Text);
        }

        //>=
        private void kryptonButton41_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton41.Text);
        }

        //<=
        private void kryptonButton42_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton42.Text);
        }

        //<>
        private void kryptonButton48_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton48.Text);
        }

        //()
        private void kryptonButton49_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton49.Text);
        }

        //{}
        private void kryptonButton50_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton50.Text);
        }

        //string
        private void kryptonButton52_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton52.Text);
        }

        //int
        private void kryptonButton53_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton53.Text);
        }

        //float
        private void kryptonButton54_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton54.Text);
        }

        //s
        private void kryptonButton71_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton71.Text);
        }

        //i
        private void kryptonButton72_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton72.Text);
        }

        //f
        private void kryptonButton73_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton73.Text);
        }

        //var
        private void kryptonButton74_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton74.Text);
        }

        //++
        private void kryptonButton69_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton69.Text);
        }

        //--
        private void kryptonButton70_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton70.Text);
        }

        //true
        private void kryptonButton43_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton43.Text);
        }

        //false
        private void kryptonButton44_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton44.Text);
        }

        //=
        private void kryptonButton56_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton56.Text);
        }

        //.
        private void kryptonButton51_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton51.Text);
        }

        //;
        private void kryptonButton55_Click(object sender, EventArgs e)
        {
            InsertSymbolOrText(kryptonButton55.Text);
        }

        private void kryptonButton20_Click(object sender, EventArgs e)
        {
            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb == null) return;

            int i = rtb.SelectionStart;
            rtb.Select(i,rtb.SelectionLength - 1);
            rtb.SelectedText = string.Empty;
        }

        private void kryptonButton75_Click(object sender, EventArgs e)
        {
            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb == null) return;

            rtb.AppendText(" ");
        }

        private void kryptonButton76_Click(object sender, EventArgs e)
        {
            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb == null) return;

            rtb.AppendText(Environment.NewLine);
        }


        private void InsertMathematicalSymbolsDialog_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Control && e.KeyCode == Keys.M)
            {
                this.Close();
            }
        }

        
        public void InsertMathematicalSymbolsDialog_FormClosing(object sender, EventArgs e)
        {
            Form1._Form1_Instance.kryptonRibbonGroupClusterButton15.Checked = false;
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                this.TopMost = true;
            }
            else
            {
                this.TopMost = false;
            }
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if(kryptonCheckBox2.Checked == true)
            {
                this.Opacity = 0.8D;
            }
            else
            {
                this.Opacity = 100;
            }
        }

        private void kryptonButton149_Click(object sender, EventArgs e)
        {
            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb == null) return;

            rtb.AppendText(currencyEdit1.DecimalValue.ToString());
        }

        private void InsertMathematicalSymbolsDialog_FormClosed(object sender, FormClosedEventArgs e)
        {

        }
    }
}
