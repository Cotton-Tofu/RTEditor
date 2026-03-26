using ComponentFactory.Krypton.Toolkit;
using Krypton.Toolkit;
using Syncfusion.Licensing;
using Syncfusion.Windows.Forms;
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
using System.Windows.Forms.Integration;

namespace RTEditor
{
    public partial class InsertCastumDateAndTimeDialog : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);
        public InsertCastumDateAndTimeDialog()
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

        private void InsertCastumDateAndTimeDialog_Load(object sender, EventArgs e)
        {
            //日付と時間の設定
            DateTime dateTime = DateTime.Now;
            kryptonNumericUpDown1.Value = dateTime.Hour;
            kryptonNumericUpDown2.Value = dateTime.Minute;
            kryptonComboBox1.Text = dateTime.ToString("tt");

            //カレンダーをロード
            CalendarReset();

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

        private void InsertCastumDateAndTimeDialog_Shown(object sender, EventArgs e)
        {

        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            SetPreviewDate();
        }

        private void clock1_ValueChanged(object sender, Syncfusion.Windows.Forms.Tools.Clock.ValueChangedEventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                DateTime dateTime = DateTime.Now;
                if (dateTime.Second == 0)
                {
                    kryptonNumericUpDown1.Value = dateTime.Hour;
                    kryptonNumericUpDown2.Value = dateTime.Minute;
                    kryptonComboBox1.Text = dateTime.ToString("tt");
                }

                kryptonNumericUpDown3.Value = dateTime.Second;
            }
        }

        private void kryptonMonthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            string w = string.Empty;
            if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Sunday)
                w = "(日曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Monday)
                w = "(月曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Tuesday)
                w = "(火曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Wednesday)
                w = "(水曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Thursday)
                w = "(木曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Friday)
                w = "(金曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Saturday)
                w = "(土曜日)";
            kryptonLabel4.Text = "選択中の日付:"+kryptonMonthCalendar1.SelectionRange.Start.Year.ToString()+"年" + kryptonMonthCalendar1.SelectionRange.Start.Month.ToString()+"月"+ kryptonMonthCalendar1.SelectionRange.Start.Day.ToString()+"日"+w;
            SetPreviewDate();
        }

        public void StartSetPreviewDate(object sender, EventArgs e)
        {
            SetPreviewDate();
        }

        public void SetPreviewDate()
        {
            //日付
            if(kryptonRadioButton1.Checked == true)
            {
                string w = string.Empty;
                if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Sunday)
                    w = "(日曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Monday)
                    w = "(月曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Tuesday)
                    w = "(火曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Wednesday)
                    w = "(水曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Thursday)
                    w = "(木曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Friday)
                    w = "(金曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Saturday)
                    w = "(土曜日)";

                kryptonLabel7.Text = kryptonMonthCalendar1.SelectionRange.Start.Year.ToString() + "年" + kryptonMonthCalendar1.SelectionRange.Start.Month.ToString() + "月" + kryptonMonthCalendar1.SelectionRange.Start.Day.ToString() + "日"+w;
            }
            //時刻
            else if(kryptonRadioButton2.Checked == true)
            {
                string h = string.Empty;
                if (kryptonNumericUpDown1.Value < 10)
                {
                    h = "0" + kryptonNumericUpDown1.Value.ToString();
                }
                else
                {
                    h = kryptonNumericUpDown1.Value.ToString();
                }

                string m = string.Empty;
                if(kryptonNumericUpDown2.Value < 10)
                {
                    m = "0"+kryptonNumericUpDown2.Value.ToString();
                }
                else
                {
                    m = kryptonNumericUpDown2.Value.ToString();
                }

                string s = string.Empty;
                if (kryptonNumericUpDown3.Value < 10)
                {
                    s = "0" + kryptonNumericUpDown3.Value.ToString();
                }
                else
                {
                    s = kryptonNumericUpDown3.Value.ToString();
                }

                kryptonLabel7.Text = kryptonComboBox1.Text+" "+h+":"+m;

                if (kryptonCheckBox2.Checked == true)
                {
                    kryptonLabel7.Text += ":"+s;
                }
            }
            //日付と時刻
            else if(kryptonRadioButton3.Checked == true)
            {
                string w = string.Empty;
                if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Sunday)
                    w = "(日曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Monday)
                    w = "(月曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Tuesday)
                    w = "(火曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Wednesday)
                    w = "(水曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Thursday)
                    w = "(木曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Friday)
                    w = "(金曜日)";
                else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Saturday)
                    w = "(土曜日)";

                string h = string.Empty;
                if (kryptonNumericUpDown1.Value < 10)
                {
                    h = "0" + kryptonNumericUpDown1.Value.ToString();
                }
                else
                {
                    h = kryptonNumericUpDown1.Value.ToString();
                }

                string m = string.Empty;
                if (kryptonNumericUpDown2.Value < 10)
                {
                    m = "0" + kryptonNumericUpDown2.Value.ToString();
                }
                else
                {
                    m = kryptonNumericUpDown2.Value.ToString();
                }

                string s = string.Empty;
                if (kryptonNumericUpDown3.Value < 10)
                {
                    s = "0" + kryptonNumericUpDown3.Value.ToString();
                }
                else
                {
                    s = kryptonNumericUpDown3.Value.ToString();
                }

                kryptonLabel7.Text = kryptonComboBox1.Text + " " + h + ":" + m;

                kryptonLabel7.Text = kryptonMonthCalendar1.SelectionRange.Start.Year.ToString() + "年" + kryptonMonthCalendar1.SelectionRange.Start.Month.ToString() + "月" + kryptonMonthCalendar1.SelectionRange.Start.Day.ToString() + "日"+w+" "+ kryptonComboBox1.Text + " " + h + ":" + m;

                if (kryptonCheckBox2.Checked == true)
                {
                    kryptonLabel7.Text += ":"+s;
                }
                    
            }
        }
        public string Date;
        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            Date = kryptonLabel7.Text;
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            //日付と時刻をすべて現在の状態にリセット
            CalendarReset();
            TimeReset();
        }

        public void CalendarReset()
        {
            //日付リセットしてから曜日取得する
            kryptonMonthCalendar1.SelectionStart = DateTime.Today;

            string w = string.Empty;
            if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Sunday)
                w = "(日曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Monday)
                w = "(月曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Tuesday)
                w = "(火曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Wednesday)
                w = "(水曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Thursday)
                w = "(木曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Friday)
                w = "(金曜日)";
            else if (kryptonMonthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Saturday)
                w = "(土曜日)";

            kryptonLabel4.Text = "選択中の日付:" + kryptonMonthCalendar1.SelectionRange.Start.Year.ToString() + "年" + kryptonMonthCalendar1.SelectionRange.Start.Month.ToString() + "月" + kryptonMonthCalendar1.SelectionRange.Start.Day.ToString() + "日"+w;
        }

        public void TimeReset()
        {
            //日付と時間の設定
            DateTime dateTime = DateTime.Now;
            kryptonNumericUpDown1.Value = dateTime.Hour;
            kryptonNumericUpDown2.Value = dateTime.Minute;
            kryptonNumericUpDown3.Value = dateTime.Second;
            kryptonComboBox1.Text = dateTime.ToString("tt");
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            SetPreviewDate();
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            CalendarReset();
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            TimeReset();
        }
    }
}
