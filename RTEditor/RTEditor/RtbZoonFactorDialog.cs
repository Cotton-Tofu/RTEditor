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
    public partial class RtbZoonFactorDialog : KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);


        public RtbZoonFactorDialog()
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

        private void trackBarEx1_Click(object sender, EventArgs e)
        {
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            Form1._Form1_Instance.kryptonTrackBar1.Value = 5;
            kryptonLabel1.Text = Form1._Form1_Instance.kryptonLabel2.Text;
        }

        private void trackBarEx1_ValueChanged(object sender, EventArgs e)
        {
            Form1._Form1_Instance.kryptonTrackBar1.Value = trackBarEx1.Value;
            kryptonLabel1.Text = Form1._Form1_Instance.kryptonLabel2.Text;
        }

        private void RtbZoonFactorDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form1._Form1_Instance.kryptonRibbonGroupButton39.Checked = false;
        }

        private void RtbZoonFactorDialog_Load(object sender, EventArgs e)
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
    }
}
