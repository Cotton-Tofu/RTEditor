using ComponentFactory.Krypton.Ribbon;
using ComponentFactory.Krypton.Toolkit;
using Microsoft.WindowsAPICodePack.Dialogs;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;
using static Syncfusion.Windows.Forms.Tools.Navigation.Bar;

namespace RTEditor
{


    public partial class SettingsDialog : KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);
        public SettingsDialog()
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

        //設定読み込みとNavigatorのスタイル変更
        private void SettingsDialog_Load(object sender, EventArgs e)
        {

            LoadSettings(sender, e);

            kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.Panel;
            kryptonNavigator1.SelectedPage = kryptonPage1;

            InstalledFontCollection fonts = new InstalledFontCollection();
            FontFamily[] fontFamilies = fonts.Families;



            foreach (FontFamily font in fontFamilies)
            {
                kryptonComboBox1.Items.Add(font.Name);
                kryptonComboBox1.AutoCompleteCustomSource.Add(font.Name);
            }

            //既定のテーマ
            if (Properties.Settings.Default.Theme == "Global")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.Global;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Global;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Global;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Global;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.Global;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Office2010;
            }
            //Professional-Systemテーマ
            else if (Properties.Settings.Default.Theme == "ProfessionalSystem")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.ProfessionalSystem;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalSystem;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalSystem;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalSystem;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalSystem;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Office2013;
            }
            //Professional-Office2003テーマ
            else if (Properties.Settings.Default.Theme == "ProfessionalOffice2003")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.ProfessionalOffice2003;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalOffice2003;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalOffice2003;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalOffice2003;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.ProfessionalOffice2003;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Office2013;
            }
            //Office2007Blueテーマ
            else if (Properties.Settings.Default.Theme == "Office2007Blue")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.Office2007Blue;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Default;
            }
            //Office2007Silverテーマ
            else if (Properties.Settings.Default.Theme == "Office2007Silver")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.Office2007Silver;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Default;
            }
            //Office2007Blackテーマ
            else if (Properties.Settings.Default.Theme == "Office2007Black")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.Office2007Black;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Default;
            }
            //Office2010Blueテーマ
            else if (Properties.Settings.Default.Theme == "Office2010Blue")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.Office2010Blue;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Office2010;
            }
            //Office2010Silverテーマ
            else if (Properties.Settings.Default.Theme == "Office2010Silver")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.Office2010Silver;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Office2010;
            }
            //Office2010Blackテーマ
            else if (Properties.Settings.Default.Theme == "Office2010Black")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.Office2010Black;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Office2010;
            }
            //SparkleBlueテーマ
            else if (Properties.Settings.Default.Theme == "SparkleBlue")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.SparkleBlue;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleBlue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleBlue;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleBlue;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleBlue;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Default;
                AllLabelsBoldPanel();
            }
            //SparkleOrangeテーマ
            else if (Properties.Settings.Default.Theme == "SparkleOrange")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.SparkleOrange;


                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleOrange;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleOrange;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleOrange;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.SparkleOrange;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Default;
                AllLabelsBoldPanel();
            }
            //SparklePurpleテーマ
            else if (Properties.Settings.Default.Theme == "SparklePurple")
            {
                kryptonPalette1.BasePaletteMode = PaletteMode.SparklePurple;

                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.SparklePurple;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.SparklePurple;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.SparklePurple;
                kryptonCommandLinkButton4.PaletteMode = Krypton.Toolkit.PaletteMode.SparklePurple;

                MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Default;
                AllLabelsBoldPanel();
            }

            if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == true)
            {
                kryptonPalette1.AllowFormChrome = InheritBool.False;
            }
            else if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == false)
            {
                kryptonPalette1.AllowFormChrome = InheritBool.True;
            }
        }

        private void AllLabelsBoldPanel()
        {

            kryptonLabel2.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel6.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel7.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel10.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel11.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel30.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel17.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel13.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel20.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel25.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel26.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            kryptonLabel33.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
        }

        private void kryptonPage5_Click(object sender, EventArgs e)
        {

        }

        //ナビゲーションバーとパネル画面連携処理
        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = true;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton9.Checked = false;
            kryptonCheckButton10.Checked = false;
            kryptonCheckButton11.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage1;
        }

        private void kryptonCheckButton2_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = true;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton9.Checked = false;
            kryptonCheckButton10.Checked = false;
            kryptonCheckButton11.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage2;
        }

        private void kryptonCheckButton3_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = true;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton9.Checked = false;
            kryptonCheckButton10.Checked = false;
            kryptonCheckButton11.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage3;
        }

        private void kryptonCheckButton4_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = true;
            kryptonCheckButton9.Checked = false;
            kryptonCheckButton10.Checked = false;
            kryptonCheckButton11.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage4;
        }

        private void kryptonCheckButton9_Click_1(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton9.Checked = true;
            kryptonCheckButton10.Checked = false;
            kryptonCheckButton11.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage5;
        }

        private void kryptonCheckButton10_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton9.Checked = false;
            kryptonCheckButton10.Checked = true;
            kryptonCheckButton11.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage6;
        }

        private void kryptonCheckButton11_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton9.Checked = false;
            kryptonCheckButton10.Checked = false;
            kryptonCheckButton11.Checked = true;

            kryptonNavigator1.SelectedPage = kryptonPage7;
        }

        //APIキーの入力を行うダイアログの表示と設定


        //APIのプロバイダーを識別する文字列型
        string APIProviderName = null;
        //ダイアログで入力したAPIキーをAPIKeyで代入させる文字列型
        string APIKey = null;

        //APIキーを保存するメソッド
        private void SaveAPIKey()
        {
            try
            {
                if (APIProviderName != null)
                {
                    if (APIProviderName == "GoogleGemini")
                    {
                        Properties.Settings.Default.GoogleGeminiAPIKey = APIKey;
                    }
                    if (APIProviderName == "GoogleOCR")
                    {
                        Properties.Settings.Default.GoogleOCRAPIKey = APIKey;
                    }
                    if (APIProviderName == "GoogleMaps")
                    {
                        Properties.Settings.Default.GoogleMapsAPIKey = APIKey;
                    }
                    if (APIProviderName == "OpenAIChatGPT")
                    {
                        Properties.Settings.Default.OpenAIChatGPTAPIKey = APIKey;
                    }

                    //各種APIキーを保存
                    Properties.Settings.Default.Save();
                    using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー保存完了", InstructionText = "APIキーを保存が完了しました", Text = "これにより外部からサービスを使用できるようになります。", OwnerWindowHandle = this.Handle}) { taskDialog.Show(); };
                }
                else
                {
                    MessageBoxAdv.Show("エラーが発生したためAPIキーは保存されませんでした。");
                }
            }
            catch
            {
                MessageBoxAdv.Show("エラーが発生したためAPIキーは保存されませんでした。");
            }
            finally
            {
                //完了後値を削除する。
                APIProviderName = null;
                APIKey = null;
            }

        }

        //APIキーを削除するメソッド
        private void DeleteAPIKey()
        {
            try
            {
                if (APIProviderName != null)
                {
                    if (APIProviderName == "GoogleGemini")
                    {
                        Properties.Settings.Default.GoogleGeminiAPIKey = string.Empty;
                    }
                    if (APIProviderName == "GoogleOCR")
                    {
                        Properties.Settings.Default.GoogleOCRAPIKey = string.Empty;
                    }
                    if (APIProviderName == "GoogleMaps")
                    {
                        Properties.Settings.Default.GoogleMapsAPIKey = string.Empty;
                    }
                    if (APIProviderName == "OpenAIChatGPT")
                    {
                        Properties.Settings.Default.OpenAIChatGPTAPIKey = string.Empty;
                    }

                    //各種APIキーを保存
                    Properties.Settings.Default.Save();
                    using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー削除完了", InstructionText = "APIキーの削除が完了しました", Text = "これにより外部からこのサービスを使用できなくなります。このデバイスに保存されているAPIキーのみを削除したため、クラウドコンソール上でのAPIキーは削除されていません。", OwnerWindowHandle = this.Handle }) { taskDialog.Show(); };
                }
                else
                {
                    MessageBoxAdv.Show("エラーが発生したためAPIキーは削除されませんでした。");
                }
            }
            catch
            {
                MessageBoxAdv.Show("エラーが発生したためAPIキーは削除されませんでした。");
            }
            finally
            {
                //完了後値を削除する。
                APIProviderName = null;
                APIKey = null;
            }
        }

        //Gemini
        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            using (Ookii.Dialogs.WinForms.InputDialog inputDialog = new Ookii.Dialogs.WinForms.InputDialog() {WindowTitle = "Google Gemini のAPIキーの設定", MainInstruction = "Google Gemini のAPIキーを入力してください", Content = "Google Gemini のAPIキーを使用することによりAI機能を RTEditor で使用することができます。APIキーの取得は Google AI Studio で入手できます。\r\n\r\nAPIキー:" })
            {
                inputDialog.Input = Properties.Settings.Default.GoogleGeminiAPIKey;
                DialogResult result = inputDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    if(inputDialog.Input != string.Empty)
                    {
                        APIProviderName = "GoogleGemini";
                        APIKey = inputDialog.Input;
                        SaveAPIKey();
                    }
                    else
                    {
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー保存失敗", InstructionText = "APIキーが空欄であるため保存できませんでした", Text = "APIキーを空欄のまま保存することはできません。APIキーをこのデバイスから削除するには「リセット」ボタンをクリックしてください。\r\nもう一度やり直してみてください。", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning, OwnerWindowHandle = this.Handle })
                        {
                            taskDialog.Show();
                            kryptonButton6_Click(sender, e);
                        };
                    }

                }
            }
        }

        //Google OCR
        private void kryptonButton11_Click(object sender, EventArgs e)
        {
            using (Ookii.Dialogs.WinForms.InputDialog inputDialog = new Ookii.Dialogs.WinForms.InputDialog() { WindowTitle = "Google OCR のAPIキーの設定", MainInstruction = "Google OCR のAPIキーを入力してください", Content = "Google OCR のAPIキーを使用することによりGoogleのクラウドを通じて高品質なOCR処理などを行えます。APIキーの取得は Google クラウドコンソール で入手できます。\r\n\r\nAPIキー:" })
            {
                inputDialog.Input = Properties.Settings.Default.GoogleOCRAPIKey;
                DialogResult result = inputDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    if (inputDialog.Input != string.Empty)
                    {
                        APIProviderName = "GoogleOCR";
                        APIKey = inputDialog.Input;
                        SaveAPIKey();
                    }
                    else
                    {
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー保存失敗", InstructionText = "APIキーが空欄であるため保存できませんでした", Text = "APIキーを空欄のまま保存することはできません。APIキーをこのデバイスから削除するには「リセット」ボタンをクリックしてください。\r\nもう一度やり直してみてください。", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning, OwnerWindowHandle = this.Handle })
                        {
                            taskDialog.Show();
                            kryptonButton11_Click(sender, e);
                        };
                    }

                }
            }
        }

        //Google マップ
        private void kryptonButton13_Click(object sender, EventArgs e)
        {
            using (Ookii.Dialogs.WinForms.InputDialog inputDialog = new Ookii.Dialogs.WinForms.InputDialog() { WindowTitle = "Google マップ のAPIキーの設定", MainInstruction = "Google マップ のAPIキーを入力してください", Content = "Google マップ のAPIキーを使用することにより地図の表示や検索などを行えます。APIキーの取得は Google クラウドコンソール で入手できます。\r\n\r\nAPIキー:" })
            {
                inputDialog.Input = Properties.Settings.Default.GoogleMapsAPIKey;
                DialogResult result = inputDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    if (inputDialog.Input != string.Empty)
                    {
                        APIProviderName = "GoogleMaps";
                        APIKey = inputDialog.Input;
                        SaveAPIKey();
                    }
                    else
                    {
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー保存失敗", InstructionText = "APIキーが空欄であるため保存できませんでした", Text = "APIキーを空欄のまま保存することはできません。APIキーをこのデバイスから削除するには「リセット」ボタンをクリックしてください。\r\nもう一度やり直してみてください。", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning, OwnerWindowHandle = this.Handle })
                        {
                            taskDialog.Show();
                            kryptonButton13_Click(sender, e);
                        };
                    }

                }
            }
        }

        //OpenAI
        private void kryptonButton9_Click(object sender, EventArgs e)
        {
            using (Ookii.Dialogs.WinForms.InputDialog inputDialog = new Ookii.Dialogs.WinForms.InputDialog() { WindowTitle = "OpenAI ChatGPT のAPIキーの設定", MainInstruction = "OpenAI ChatGPT のAPIキーを入力してください", Content = "OpenAI ChatGPT のAPIキーを使用することによりAI機能を RTEditor で使用することができます。APIキーの取得は OpenAI プラットフォーム で入手できます。\r\n\r\nAPIキー:" })
            {
                inputDialog.Input = Properties.Settings.Default.OpenAIChatGPTAPIKey;
                DialogResult result = inputDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    if (inputDialog.Input != string.Empty)
                    {
                        APIProviderName = "OpenAIChatGPT";
                        APIKey = inputDialog.Input;
                        SaveAPIKey();
                    }
                    else
                    {
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー保存失敗", InstructionText = "APIキーが空欄であるため保存できませんでした", Text = "APIキーを空欄のまま保存することはできません。APIキーをこのデバイスから削除するには「リセット」ボタンをクリックしてください。\r\nもう一度やり直してみてください。", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning, OwnerWindowHandle = this.Handle }) 
                        { 
                            taskDialog.Show();
                            kryptonButton9_Click(sender, e);
                        };
                    }

                }
            }
        }

        //APIキーをリセットする処理

        //Gemini
        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー削除確認", InstructionText = "Google Gemini のAPIキーを削除しますか?", Text = "Google Gemini のAPIキー完全に削除しようとしています。この操作を行うと Google Gemini のAI機能を RTEditor で使用できなくなります。このデバイスに保存されているAPIキーを削除しますがクラウドコンソール上に保存されているAPIキーは削除されません。本当によろしいですか?", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Information, StandardButtons = TaskDialogStandardButtons.Yes|TaskDialogStandardButtons.No,OwnerWindowHandle = this.Handle })
            {
                TaskDialogResult tskDialogResult = taskDialog.Show();
                if(tskDialogResult == TaskDialogResult.Yes)
                {
                    APIProviderName = "GoogleGemini";
                    DeleteAPIKey();
                }
            }
        }

        //Google OCR
        private void kryptonButton10_Click(object sender, EventArgs e)
        {
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー削除確認", InstructionText = "Google OCR のAPIキーを削除しますか?", Text = "Google OCR のAPIキー完全に削除しようとしています。この操作を行うとクラウド上でのOCR機能を RTEditor で使用できなくなります。このデバイスに保存されているAPIキーを削除しますがクラウドコンソール上に保存されているAPIキーは削除されません。本当によろしいですか?", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Information, StandardButtons = TaskDialogStandardButtons.Yes | TaskDialogStandardButtons.No, OwnerWindowHandle = this.Handle })
            {
                TaskDialogResult tskDialogResult = taskDialog.Show();
                if (tskDialogResult == TaskDialogResult.Yes)
                {
                    APIProviderName = "GoogleOCR";
                    DeleteAPIKey();
                }
            }
        }

        //Google マップ
        private void kryptonButton12_Click(object sender, EventArgs e)
        {
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー削除確認", InstructionText = "Google マップ のAPIキーを削除しますか?", Text = "Google マップ のAPIキー完全に削除しようとしています。この操作を行うと Google マップ の表示や地図の検索等を RTEditor で使用できなくなります。このデバイスに保存されているAPIキーを削除しますがクラウドコンソール上に保存されているAPIキーは削除されません。本当によろしいですか?", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Information, StandardButtons = TaskDialogStandardButtons.Yes | TaskDialogStandardButtons.No, OwnerWindowHandle = this.Handle })
            {
                TaskDialogResult tskDialogResult = taskDialog.Show();
                if (tskDialogResult == TaskDialogResult.Yes)
                {
                    APIProviderName = "GoogleMaps";
                    DeleteAPIKey();
                }
            }
        }

        //OpenAI ChatGPT
        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー削除確認", InstructionText = "OpenAI ChatGPT のAPIキーを削除しますか?", Text = "OpenAI ChatGPT のAPIキー完全に削除しようとしています。この操作を行うと OpenAI ChatGPT のAI機能を RTEditor で使用できなくなります。このデバイスに保存されているAPIキーを削除しますがクラウドコンソール上に保存されているAPIキーは削除されません。本当によろしいですか?", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Information, StandardButtons = TaskDialogStandardButtons.Yes | TaskDialogStandardButtons.No, OwnerWindowHandle = this.Handle })
            {
                TaskDialogResult tskDialogResult = taskDialog.Show();
                if (tskDialogResult == TaskDialogResult.Yes)
                {
                    APIProviderName = "OpenAIChatGPT";
                    DeleteAPIKey();
                }
            }
        }

        //全サービスのAPIキー削除
        private void kryptonButton14_Click(object sender, EventArgs e)
        {
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー削除確認", InstructionText = "すべてのサービスのAPIキーを削除しますか?", Text = "すべてのサービスのAPIキー完全に削除しようとしています。この操作を行うとすべてのサービスのAPIキーがこのデバイスのみ削除され Google や OpenAI 製の機能をRTEditorで使用できなくなります。このデバイスに保存されているすべてのサービスのAPIキーは削除されますが、クラウドコンソール上の各種サービスのAPIキーは削除されません。本当によろしいですか?", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning, StandardButtons = TaskDialogStandardButtons.Yes | TaskDialogStandardButtons.No, OwnerWindowHandle = this.Handle })
            {
                TaskDialogResult tskDialogResult = taskDialog.Show();
                if (tskDialogResult == TaskDialogResult.Yes)
                {
                    try
                    {
                        Properties.Settings.Default.GoogleGeminiAPIKey = string.Empty;
                        Properties.Settings.Default.GoogleOCRAPIKey = string.Empty;
                        Properties.Settings.Default.GoogleMapsAPIKey = string.Empty;
                        Properties.Settings.Default.OpenAIChatGPTAPIKey = string.Empty;
                        Properties.Settings.Default.Save();
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog2 = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "APIキー全削除完了", InstructionText = "APIキーの全削除が完了しました", Text = "これにより外部からすべてのサービスを使用できなくなります。このデバイスに保存されているAPIキーのみをすべて削除したため、クラウドコンソール上の各種サービスでのAPIキーは削除されていません。", OwnerWindowHandle = this.Handle }) { taskDialog2.Show(); };
                    }
                    catch
                    {
                        MessageBoxAdv.Show("エラーが発生したためAPIキーは削除されませんでした。");
                    }

                }
            }
        }


        //ダイアログとポップアップのリセット
        private void commandLink1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.IsShowDeleteWarningTaskDialog = true;
            Properties.Settings.Default.Save();
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "完了", InstructionText = "ダイアログやポップアップ等の表示設定をリセットしました", Text = "これにより再度必要に応じてダイアログやポップアップ等が表示されるようになります。", OwnerWindowHandle = this.Handle }) { taskDialog.Show(); };
        }

        //全設定のリセット
        private void commandLink2_Click(object sender, EventArgs e)
        {
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "設定リセット確認", InstructionText = "RTEditor の設定をすべてリセットします", Text = "RTEditor で変更した設定を初回起動時の状態にすべてリセットします。また全サービスのAPIキーや最近使用したドキュメントの項目をデバイスから削除します。設定をすべてリセットした場合、元に戻すことができません。本当に設定をすべてリセットしますか?", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning, StandardButtons = TaskDialogStandardButtons.Yes | TaskDialogStandardButtons.No, OwnerWindowHandle = this.Handle })
            {
                TaskDialogResult tskDialogResult = taskDialog.Show();
                if (tskDialogResult == TaskDialogResult.Yes)
                {
                    try
                    {
                        //設定をすべてリセットする
                        Properties.Settings.Default.Reset();
                        Properties.Settings.Default.Save();
                        LoadSettings(sender, e);
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog DoneDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "完了", InstructionText = "設定をすべてリセットしました", Text = "これによりすべての設定が初期値にリセットされました。", OwnerWindowHandle = this.Handle }) { DoneDialog.Show(); }
                        ;
                    }
                    catch
                    {
                    }

                }
            }
        }

        //OKボタンが押されたときの設定の保存処理
        private void kryptonButton2_Click(object sender, EventArgs e)
        {

            //設定保存処理

            //一般タブ

            //ユーザーインターフェースのオプション

            //ミニツールバーとラジアルメニュー
            if (kryptonRadioButton1.Checked == true)
            {
                Properties.Settings.Default.MiniToolBarOrRadialMenu = "MiniToolBar";
            }
            else if (kryptonRadioButton2.Checked == true)
            {
                Properties.Settings.Default.MiniToolBarOrRadialMenu = "RadialMenu";
            }
            else if(kryptonRadioButton5.Checked == true)
            {
                Properties.Settings.Default.MiniToolBarOrRadialMenu = "None";
            }

            //リボンのタッチ操作向けUI調整
            if (kryptonCheckBox1.Checked == true)
            {
                Properties.Settings.Default.IsOptimizedTachModeRibbon = true;
            }
            else if (kryptonCheckBox1.Checked == false)
            {
                Properties.Settings.Default.IsOptimizedTachModeRibbon = false;
            }

            //テーマ
            PaletteMode paletteMode = (PaletteMode)kryptonComboBox4.SelectedIndex;
            //テーマの名前とコンボボックスの値を保存する
            Properties.Settings.Default.Theme = paletteMode.ToString();
            Properties.Settings.Default.Theme_IntType = (int)kryptonComboBox4.SelectedIndex;

            //システムスタイルのタイトルバー
            if(kryptonCheckBox7.Checked == true)
            {
                Properties.Settings.Default.IsUseSystemTitleBar = true;
            }
            else if(kryptonCheckBox7.Checked == false)
            {
                Properties.Settings.Default.IsUseSystemTitleBar = false;
            }

            //ダイアログやMdiウィンドウのシステムスタイルのタイトルバー使用
            if (kryptonCheckBox12.Checked == true)
            {
                Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar = true;
            }
            else if (kryptonCheckBox12.Checked == false)
            {
                Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar = false;
            }

            //タイトルバーのダークモード
            if (kryptonCheckBox14.Checked == true)
            {
                Properties.Settings.Default.IsUseDarkModeTitleBar = true;

            }
            else if (kryptonCheckBox14.Checked == false)
            {
                Properties.Settings.Default.IsUseDarkModeTitleBar = false;
            }

            //Office 2007のリボンシェイプ
            if (kryptonCheckBox8.Checked == true)
            {
                Properties.Settings.Default.IsUseOffice2007RibbonShape = true;
            }
            else if (kryptonCheckBox8.Checked == false)
            {
                Properties.Settings.Default.IsUseOffice2007RibbonShape = false;
            }

            //起動時の設定

            if (kryptonRadioButton3.Checked == true)
            {
                //何もしない
                Properties.Settings.Default.AppStartupTaskMode = 0;
            }
            else if (kryptonRadioButton4.Checked == true)
            {
                //新しいドキュメントを自動作成
                Properties.Settings.Default.AppStartupTaskMode = 1;
            }

            //既定のフォントとスタイル

            //既定のフォント名
            Properties.Settings.Default.DefaultFontName = kryptonComboBox1.Text;
            //既定のフォントサイズ(フォントサイズの設定に使うのでTryParseで変換)
            if (float.TryParse(kryptonComboBox2.Text, out float fontSize))
            {
                Properties.Settings.Default.DefaultFontSize = fontSize;
            }

            //太字
            if(kryptonCheckButton5.Checked == true)
            {
                Properties.Settings.Default.DefaultFontIsBold = true;
            }
            else if(kryptonCheckButton5.Checked == false)
            {
                Properties.Settings.Default.DefaultFontIsBold = false;
            }

            //斜体
            if (kryptonCheckButton6.Checked == true)
            {
                Properties.Settings.Default.DefaultFontIsItalic = true;
            }
            else if (kryptonCheckButton6.Checked == false)
            {
                Properties.Settings.Default.DefaultFontIsItalic = false;
            }

            //下線
            if (kryptonCheckButton7.Checked == true)
            {
                Properties.Settings.Default.DefaultFontIsUnderline = true;
            }
            else if (kryptonCheckButton7.Checked == false)
            {
                Properties.Settings.Default.DefaultFontIsUnderline = false;
            }

            //打ち消し線
            if (kryptonCheckButton8.Checked == true)
            {
                Properties.Settings.Default.DefaultFontIsStrikeout = true;
            }
            else if (kryptonCheckButton8.Checked == false)
            {
                Properties.Settings.Default.DefaultFontIsStrikeout = false;
            }
            //文字色
            Properties.Settings.Default.DefaultFontTextColor = kryptonColorButton1.SelectedColor;
            //蛍光ペン
            Properties.Settings.Default.DefaultFontTextHilightColor = kryptonColorButton2.SelectedColor;


            //表示タブ

            //ドキュメントの表示オプション

            //既定でドキュメント内の文字列の折り返し
            if (kryptonCheckBox2.Checked == true)
            {
                Properties.Settings.Default.IsUseWordWarp = true;
            }
            else if (kryptonCheckBox2.Checked == false)
            {
                Properties.Settings.Default.IsUseWordWarp = false;
            }

            //左端の段落選択用の空白の表示
            if (kryptonCheckBox3.Checked == true)
            {
                Properties.Settings.Default.IsShowSelectionMargin = true;
            }
            else if (kryptonCheckBox3.Checked == false)
            {
                Properties.Settings.Default.IsShowSelectionMargin = false;
            }

            //単語自動選択機能
            if (kryptonCheckBox9.Checked == true)
            {
                Properties.Settings.Default.IsUseAutoWordSelection= true;
            }
            else if (kryptonCheckBox9.Checked == false)
            {
                Properties.Settings.Default.IsUseAutoWordSelection = false;
            }

            //リッチテキストエディタのダークモード
            if (kryptonCheckBox13.Checked == true)
            {
                Properties.Settings.Default.IsUseRichTextBoxDarkMode = true;
            }
            else if (kryptonCheckBox13.Checked == false)
            {
                Properties.Settings.Default.IsUseRichTextBoxDarkMode = false;
            }

            //AI機能のオプション
            //AI機能の表示
            if (kryptonCheckBox4.Checked == true)
            {
                Properties.Settings.Default.IsUseAIFunction = true;
            }
            else if (kryptonCheckBox4.Checked == false)
            {
                Properties.Settings.Default.IsUseAIFunction = false;
            }

            //その他のオプション
            //ステータスバーの表示
            if (kryptonCheckBox10.Checked == true)
            {
                Properties.Settings.Default.IsShowStatusBar = true;
            }
            else if (kryptonCheckBox10.Checked == false)
            {
                Properties.Settings.Default.IsShowStatusBar = false;
            }
            //既定でリボンを最小化
            if (kryptonCheckBox11.Checked == true)
            {
                Properties.Settings.Default.IsMinimizedRibbon = true;
            }
            else if (kryptonCheckBox11.Checked == false)
            {
                Properties.Settings.Default.IsMinimizedRibbon = false;
            }

            //保存タブ
            //ドキュメントの保存
            //既定のファイル保存形式(int型で保存)
            //0でリッチテキスト、1で書式なしテキスト、
            Properties.Settings.Default.DefaultSaveMode = kryptonComboBox3.SelectedIndex;
            //自動保存機能の有効・無効
            if (kryptonCheckBox5.Checked == true)
            {
                Properties.Settings.Default.IsUseAutoSave = true;
            }
            else if (kryptonCheckBox5.Checked == false)
            {
                Properties.Settings.Default.IsUseAutoSave = false;
            }
            //自動保存を行う間隔(整数×60000でミリ秒単位に換算、割り算で分に換算する)
            Properties.Settings.Default.AutoSaveTime = (int)kryptonNumericUpDown1.Value*60000;

            //ファイルの関連付け(なし)
            //APIキー(なし)
            //リセット(なし)

            //クイックアクセスツールバー
            string result = "";
            foreach (string item in kryptonCheckedListBox1.CheckedItems)
            {
                result += item+Environment.NewLine;
            }

            if (result.Contains("新しいドキュメント") == true)
            {
                Properties.Settings.Default.QAT1Visible = true;
            }
            else if (result.Contains("新しいドキュメント") == false)
            {
                Properties.Settings.Default.QAT1Visible = false;
            }

            if (result.Contains("開く") == true)
            {
                Properties.Settings.Default.QAT2Visible = true;
            }
            else if (result.Contains("開く") == false)
            {
                Properties.Settings.Default.QAT2Visible = false;
            }

            if (result.Contains("保存 ") == true)
            {
                Properties.Settings.Default.QAT3Visible = true;
            }
            else if (result.Contains("保存 ") == false)
            {
                Properties.Settings.Default.QAT3Visible = false;
            }

            if (result.Contains("名前を付けて保存") == true)
            {
                Properties.Settings.Default.QAT4Visible = true;
            }
            else if (result.Contains("名前を付けて保存") == false)
            {
                Properties.Settings.Default.QAT4Visible = false;
            }

            if (result.Contains("印刷 ") == true)
            {
                Properties.Settings.Default.QAT5Visible = true;
            }
            else if (result.Contains("印刷 ") == false)
            {
                Properties.Settings.Default.QAT5Visible = false;
            }

            if (result.Contains("印刷プレビュー") == true)
            {
                Properties.Settings.Default.QAT6Visible = true;
            }
            else if (result.Contains("印刷プレビュー") == false)
            {
                Properties.Settings.Default.QAT6Visible = false;
            }

            if (result.Contains("元に戻す") == true)
            {
                Properties.Settings.Default.QAT7Visible = true;
            }
            else if (result.Contains("元に戻す") == false)
            {
                Properties.Settings.Default.QAT7Visible = false;
            }

            if (result.Contains("やり直す") == true)
            {
                Properties.Settings.Default.QAT8Visible = true;
            }
            else if (result.Contains("やり直す") == false)
            {
                Properties.Settings.Default.QAT8Visible = false;
            }

            //設定の保存
            Properties.Settings.Default.Save();

        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
        }

        //すべての拡張子を選択するイベントハンドラ
        private void kryptonButton19_Click(object sender, EventArgs e)
        {
            int i = 0;
            while(i < kryptonCheckedListBox2.Items.Count)
            {
                kryptonCheckedListBox2.SetItemChecked(i, true);
                i++;
                if(i == kryptonCheckedListBox2.Items.Count)
                {
                    break;
                }

            }
        }

        //すべての拡張子を非選択するイベントハンドラ
        private void kryptonButton20_Click(object sender, EventArgs e)
        {
            int i = 0;
            while (i < kryptonCheckedListBox2.Items.Count)
            {
                kryptonCheckedListBox2.SetItemChecked(i, false);
                i++;
                if (i == kryptonCheckedListBox2.Items.Count)
                {
                    break;
                }

            }
        }

        //関連付けを行うイベントハンドラ
        private void kryptonButton5_Click(object sender, EventArgs e)
        {

        }

        //関連付けを解除するイベントハンドラ
        private void kryptonButton4_Click(object sender, EventArgs e)
        {

        }


        public void LoadSettings(object sender, EventArgs e)
        {
            //ユーザーインターフェースのオプション

            //ミニツールバーとラジアルメニュー
            if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "MiniToolBar")
            {
                kryptonRadioButton1.Checked = true;
            }
            else if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "RadialMenu")
            {
                kryptonRadioButton2.Checked = true;
            }
            else if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "None")
            {
                kryptonRadioButton5.Checked = true;
            }


            //リボンのタッチ操作向けUI調整
            if (Properties.Settings.Default.IsOptimizedTachModeRibbon == true)
            {
                kryptonCheckBox1.Checked = true;
            }
            else if (Properties.Settings.Default.IsOptimizedTachModeRibbon == false)
            {
                kryptonCheckBox1.Checked = false;
            }

            //テーマ
            //以前保存したコンボボックスの値をもとにテーマを割り出す
            kryptonComboBox4.SelectedIndex = (int)Properties.Settings.Default.Theme_IntType;

            //システムスタイルのタイトルバー
            if (Properties.Settings.Default.IsUseSystemTitleBar == true)
            {
                kryptonCheckBox7.Checked = true;
            }
            else if (Properties.Settings.Default.IsUseSystemTitleBar == false)
            {
                kryptonCheckBox7.Checked = false;
            }

            //ダイアログやMdiウィンドウのシステムスタイルのタイトルバー使用
            if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == true)
            {
                kryptonCheckBox12.Checked = true;
            }
            else if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == false)
            {
                kryptonCheckBox12.Checked = false;
            }

            //タイトルバーのダークモード
            if (Properties.Settings.Default.IsUseDarkModeTitleBar == true)
            {
                kryptonCheckBox14.Checked = true;

            }
            else if (Properties.Settings.Default.IsUseDarkModeTitleBar == false)
            {
                kryptonCheckBox14.Checked = false;
            }

            //リボンシェイプ
            if (Properties.Settings.Default.IsUseOffice2007RibbonShape == true)
            {
                kryptonCheckBox8.Checked = true;
            }
            else if (Properties.Settings.Default.IsUseOffice2007RibbonShape == false)
            {
                kryptonCheckBox8.Checked = false;
            }


            //起動時の設定

            if (Properties.Settings.Default.AppStartupTaskMode == 0)
            {
                //何もしない
                kryptonRadioButton3.Checked = true;
            }
            else if (Properties.Settings.Default.AppStartupTaskMode == 1)
            {
                //新しいドキュメントを自動作成
                kryptonRadioButton4.Checked = true;
            }

            //既定のフォントとスタイル

            //既定のフォント名
            kryptonComboBox1.Text = Properties.Settings.Default.DefaultFontName;
            //既定のフォントサイズ
            kryptonComboBox2.Text = Properties.Settings.Default.DefaultFontSize.ToString();

            //太字
            if (Properties.Settings.Default.DefaultFontIsBold == true)
            {
                kryptonCheckButton5.Checked = true;
            }
            else if (Properties.Settings.Default.DefaultFontIsBold == false)
            {
                
                kryptonCheckButton5.Checked = false;
            }

            //斜体
            if (Properties.Settings.Default.DefaultFontIsItalic == true)
            {
                kryptonCheckButton6.Checked = true;
            }
            else if (Properties.Settings.Default.DefaultFontIsItalic == false)
            {
                kryptonCheckButton6.Checked = false;
            }

            //下線
            if (Properties.Settings.Default.DefaultFontIsUnderline == true)
            {
                kryptonCheckButton7.Checked = true;
            }
            else if (Properties.Settings.Default.DefaultFontIsUnderline == false)
            {
                kryptonCheckButton7.Checked = false;
            }

            //打ち消し線
            if (Properties.Settings.Default.DefaultFontIsStrikeout == true)
            {
                kryptonCheckButton8.Checked = true;
            }
            else if (Properties.Settings.Default.DefaultFontIsStrikeout == false)
            {
                kryptonCheckButton8.Checked = false;
            }

            //フォントの適用
            label1.Font = new Font(Properties.Settings.Default.DefaultFontName, Properties.Settings.Default.DefaultFontSize, label1.Font.Style);
            //各フォントスタイルの適用
            kryptonCheckButton5_Click(sender, e);
            kryptonCheckButton6_Click(sender, e);
            kryptonCheckButton7_Click(sender, e);
            kryptonCheckButton8_Click(sender, e);

            //文字色
            kryptonColorButton1.SelectedColor = Properties.Settings.Default.DefaultFontTextColor;
            //蛍光ペン
            kryptonColorButton2.SelectedColor = Properties.Settings.Default.DefaultFontTextHilightColor;

            //文字色と蛍光ペンの適用
            label1.ForeColor = Properties.Settings.Default.DefaultFontTextColor;
            label1.BackColor = Properties.Settings.Default.DefaultFontTextHilightColor;


            //表示タブ

            //ドキュメントの表示オプション

            //既定でドキュメント内の文字列の折り返し
            if (Properties.Settings.Default.IsUseWordWarp == true)
            {
                kryptonCheckBox2.Checked = true;
            }
            else if (Properties.Settings.Default.IsUseWordWarp == false)
            {
                kryptonCheckBox2.Checked = false;
            }

            //左端の段落選択用の空白の表示
            if (Properties.Settings.Default.IsShowSelectionMargin == true)
            {
                kryptonCheckBox3.Checked = true;
            }
            else if (Properties.Settings.Default.IsShowSelectionMargin == false)
            {
                kryptonCheckBox3.Checked = false;
            }

            //単語自動選択機能
            if (Properties.Settings.Default.IsUseAutoWordSelection == true)
            {
                kryptonCheckBox9.Checked = true;
            }
            else if (Properties.Settings.Default.IsUseAutoWordSelection == false)
            {
                kryptonCheckBox9.Checked = false;
            }

            //リッチテキストエディタのダークモード
            if(Properties.Settings.Default.IsUseRichTextBoxDarkMode == true)
            {
                kryptonCheckBox13.Checked = true;
            }
            else if(Properties.Settings.Default.IsUseRichTextBoxDarkMode == false)
            {
                kryptonCheckBox13.Checked = false;
            }

            //AI機能のオプション
            //AI機能の表示
            if (Properties.Settings.Default.IsUseAIFunction == true)
            {
                kryptonCheckBox4.Checked = true;
            }
            else if (Properties.Settings.Default.IsUseAIFunction == false)
            {
                kryptonCheckBox4.Checked = false;
            }

            //その他のオプション
            //ステータスバーの表示
            if (Properties.Settings.Default.IsShowStatusBar == true)
            {
                kryptonCheckBox10.Checked = true;
            }
            else if (Properties.Settings.Default.IsShowStatusBar == false)
            {
                kryptonCheckBox10.Checked = false;
            }
            //既定でリボンを最小化
            if (Properties.Settings.Default.IsMinimizedRibbon == true)
            {
                kryptonCheckBox11.Checked = true;
            }
            else if (Properties.Settings.Default.IsMinimizedRibbon == false)
            {
                kryptonCheckBox11.Checked = false;
            }

            //保存タブ
            //ドキュメントの保存
            //既定のファイル保存形式(int型で保存)
            //0でリッチテキスト、1で書式なしテキスト、
            kryptonComboBox3.SelectedIndex = Properties.Settings.Default.DefaultSaveMode;
            //自動保存機能の有効・無効
            if (Properties.Settings.Default.IsUseAutoSave == true)
            {
                kryptonCheckBox5.Checked = true;
            }
            else if (Properties.Settings.Default.IsUseAutoSave == false)
            {
                kryptonCheckBox5.Checked = false;
            }

            kryptonCheckBox5_CheckedChanged(sender, e);
            //自動保存を行う間隔(整数×60000でミリ秒単位に換算、割り算で分に換算する)
             kryptonNumericUpDown1.Value = (int)Properties.Settings.Default.AutoSaveTime / 60000;

            //ファイルの関連付け(なし)
            //APIキー(なし)
            //リセット(なし)

            //クイックアクセスツールバー

            if (Properties.Settings.Default.QAT1Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(0, true);
            }
            if (Properties.Settings.Default.QAT2Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(1, true);
            }
            if (Properties.Settings.Default.QAT3Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(2, true);
            }
            if (Properties.Settings.Default.QAT4Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(3, true);
            }
            if (Properties.Settings.Default.QAT5Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(4, true);
            }
            if (Properties.Settings.Default.QAT6Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(5, true);
            }
            if (Properties.Settings.Default.QAT7Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(6, true);
            }
            if (Properties.Settings.Default.QAT8Visible == true)
            {
                kryptonCheckedListBox1.SetItemChecked(7, true);
            }



        }

        private void kryptonComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void SettingsDialog_Shown(object sender, EventArgs e)
        {
        }



        //フォントスタイルに関する処理

        //フォントのリセット
        public void FontStyleReset()
        {
            label1.Font = new Font(label1.Font.Name, label1.Font.Size, FontStyle.Regular);
        }

        //太字が有効な場合
        private void kryptonCheckButton5_Click(object sender, EventArgs e)
        {
            if(kryptonCheckButton5.Checked == true)
            {
                label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style|FontStyle.Bold);
            }
            else
            {
                FontStyleReset();
                //斜体が有効な場合
                if (kryptonCheckButton6.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Italic);
                }

                //下線が有効な場合
                if (kryptonCheckButton7.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Underline);
                }

                //打ち消し線が有効な場合
                if (kryptonCheckButton8.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Strikeout);
                }
            }
        }

        //斜体が有効な場合
        private void kryptonCheckButton6_Click(object sender, EventArgs e)
        {
            if (kryptonCheckButton6.Checked == true)
            {
                label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Italic);
            }
            else
            {
                FontStyleReset();
                //太字が有効な場合
                if (kryptonCheckButton5.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Bold);
                }

                //下線が有効な場合
                if (kryptonCheckButton7.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Underline);
                }

                //打ち消し線が有効な場合
                if (kryptonCheckButton8.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Strikeout);
                }
            }
        }

        //下線が有効な場合
        private void kryptonCheckButton7_Click(object sender, EventArgs e)
        {
            if (kryptonCheckButton7.Checked == true)
            {
                label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Underline);
            }
            else
            {
                FontStyleReset();
                //太字が有効な場合
                if (kryptonCheckButton5.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Bold);
                }

                //斜体が有効な場合
                if (kryptonCheckButton6.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Italic);
                }

                //打ち消し線が有効な場合
                if (kryptonCheckButton8.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Strikeout);
                }
            }
        }

        //打ち消し線が有効な場合
        private void kryptonCheckButton8_Click(object sender, EventArgs e)
        {
            if (kryptonCheckButton8.Checked == true)
            {
                label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Strikeout);
            }
            else
            {
                FontStyleReset();
                //太字が有効な場合
                if (kryptonCheckButton5.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Bold);
                }

                //斜体が有効な場合
                if (kryptonCheckButton6.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Italic);
                }

                //下線が有効な場合
                if (kryptonCheckButton7.Checked == true)
                {
                    label1.Font = new Font(label1.Font.Name, label1.Font.Size, label1.Font.Style | FontStyle.Underline);
                }
            }
        }

        private void kryptonPanel13_Paint(object sender, PaintEventArgs e)
        {

        }

        //文字色変更
        private void kryptonColorButton1_SelectedColorChanged(object sender, ColorEventArgs e)
        {
            label1.ForeColor = e.Color;
        }

        //蛍光ペン色変更
        private void kryptonColorButton2_SelectedColorChanged(object sender, ColorEventArgs e)
        {
            label1.BackColor = e.Color;
        }

        //フォントの変更
        private void kryptonComboBox1_TextUpdate(object sender, EventArgs e)
        {
            try
            {
                label1.Font = new Font(kryptonComboBox1.Text, label1.Font.Size, label1.Font.Style);
            }
            catch
            { }
        }

        //フォントサイズ変更
        private void kryptonComboBox2_TextUpdate(object sender, EventArgs e)
        {
            if(float.TryParse(kryptonComboBox2.Text, out float fontsize))
            {
                label1.Font = new Font(label1.Font.Name, fontsize, label1.Font.Style);
            }
        }

        //フォントダイアログを表示してフォントの変更を行う
        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            using (FontDialog fontDialog = new FontDialog() { Font = label1.Font, ShowColor = true, Color = label1.ForeColor})
            {
                if(fontDialog.ShowDialog() == DialogResult.OK)
                {
                    label1.Font = new Font(fontDialog.Font.Name, fontDialog.Font.Size, fontDialog.Font.Style);
                    label1.ForeColor = fontDialog.Color;

                    kryptonComboBox1.Text = fontDialog.Font.Name;
                    kryptonComboBox2.Text = fontDialog.Font.Size.ToString();

                    if(fontDialog.Font.Bold)
                    {
                        kryptonCheckButton5.Checked = true;
                    }
                    else
                    {
                        kryptonCheckButton5.Checked = false;
                    }

                    if (fontDialog.Font.Italic)
                    {
                        kryptonCheckButton6.Checked = true;
                    }
                    else
                    {
                        kryptonCheckButton6.Checked = false;
                    }

                    if (fontDialog.Font.Underline)
                    {
                        kryptonCheckButton7.Checked = true;
                    }
                    else
                    {
                        kryptonCheckButton7.Checked = false;
                    }

                    if (fontDialog.Font.Strikeout)
                    {
                        kryptonCheckButton8.Checked = true;
                    }
                    else
                    {
                        kryptonCheckButton8.Checked = false;
                    }

                    kryptonColorButton1.SelectedColor = fontDialog.Color;
                }
            }
        }

        //タイマーによる自動保存の有効・無効による自動保存間隔の設定の有効・無効の切り替え
        private void kryptonCheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            if(kryptonCheckBox5.Checked == true)
            {
                kryptonNumericUpDown1.Enabled = true;
                kryptonLabel15.Enabled = true;
            }
            else if(kryptonCheckBox5.Checked == false)
            {
                kryptonNumericUpDown1.Enabled = false;
                kryptonLabel15.Enabled = false;
            }
        }

        private void kryptonButton21_Click(object sender, EventArgs e)
        {

        }

        //最近使用したドキュメントの一覧を削除
        private void kryptonCommandLinkButton3_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.RecentDocsPath = string.Empty;
            Properties.Settings.Default.Save();
            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "完了", InstructionText = "最近使用したドキュメントの一覧をすべて削除しました",Text = "最近使用したドキュメントの一覧はすべて削除しましたが再度ファイルの読み書きを行うと自動的に追加されます。", OwnerWindowHandle = this.Handle }) { taskDialog.Show(); }
            ;
        }

        //キャッシュファイルの削除
        private void kryptonCommandLinkButton4_Click(object sender, EventArgs e)
        {
            using(Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog(){Caption = "キャッシュファイルの削除", InstructionText = "Edge WebView2とRTEditorの設定保存用キャッシュファイルを削除します", Text = "この操作を続行するとEdge WebView2のランタイムとRTEditorの設定が削除されアプリが初期化されこのソフトウェアを閉じます。閉じる前にすべてのファイルを保存しておくことを推奨します。この操作は元に戻せません。よろしいですか?",Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Information, StandardButtons = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardButtons.Yes | Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardButtons.No, OwnerWindowHandle = this.Handle })
            {

                if(taskDialog.Show() == TaskDialogResult.Yes)
                {
                    try
                    {
                        //削除確認のメッセージボックスを表示
                        string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

                        // キャッシュファイルの削除処理
                        string edgeWebView2CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\RTEditor");
                        string rtEditorCachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Cotton_Tofu");

                        if (System.IO.Directory.Exists(edgeWebView2CachePath))
                        {
                            System.IO.Directory.Delete(edgeWebView2CachePath, true);
                        }
                        if (System.IO.Directory.Exists(rtEditorCachePath))
                        {
                            System.IO.Directory.Delete(rtEditorCachePath, true);
                        }

                        Application.Exit();
                    }
                    catch (Exception ex)
                    {
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog ErrDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "キャッシュファイルの削除中にエラーが発生しました", Text = $"エラーの詳細: {ex.Message}", OwnerWindowHandle = this.Handle }) { ErrDialog.Show(); };
                    }
                }
            }
        }
    }
}
