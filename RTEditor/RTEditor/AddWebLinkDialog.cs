using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RTEditor
{
    public partial class AddWebLinkDialog : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);

        public AddWebLinkDialog()
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

        public string WebLinkURL { get; set; }
        private void AddWebLinkDialog_Load(object sender, EventArgs e)
        {
            webView21.Source = new Uri("about:blank");

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

        private void webView21_ContentLoading(object sender, Microsoft.Web.WebView2.Core.CoreWebView2ContentLoadingEventArgs e)
        {
            if(webView21.Source != new Uri("about:blank"))
            {
                buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                //常にkryptonButton1のURLと同じページを表示する
                kryptonButton1_Click(sender, e);
            }
            else
            {
                buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
            }
            kryptonTextBox2.Text = webView21.Source.ToString();
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox1.Text != string.Empty)
            {
                try
                {
                    if (kryptonTextBox1.Text.Contains("https") == true | kryptonTextBox1.Text.Contains("http") == true)
                    {
                        Uri uri = new Uri(kryptonTextBox1.Text);
                        webView21.Source = uri;
                        kryptonLabel6.Visible = false;
                    }
                    else
                    {
                        Uri uri = new Uri("https://" + kryptonTextBox1.Text);
                            
                        webView21.Source = uri;
                        kryptonLabel6.Visible = false;

                    }

                }
                catch
                {
                    kryptonLabel6.Visible = true;
                }

            }
            else
            {
                webView21.Source = new Uri("about:blank");
            }

        }

        private void webView21_NavigationStarting(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs e)
        {

        }

        private void webView21_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
        }

        private void buttonSpecAny1_Click(object sender, EventArgs e)
        {
            kryptonTextBox1.Text = string.Empty;
            webView21.Source = new Uri("about:blank");
            kryptonLabel6.Visible = false;
        }

        private void buttonSpecAny4_Click(object sender, EventArgs e)
        {
            Process.Start(webView21.Source.ToString());
        }

        private void kryptonTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyValue == (int)Keys.Enter)
            {
                kryptonButton1_Click(sender, e);
            }
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                WebLinkURL = webView21.Source.ToString();
            }
            else
            {
                WebLinkURL = kryptonTextBox1.Text;
            }
            
        }
    }
}
