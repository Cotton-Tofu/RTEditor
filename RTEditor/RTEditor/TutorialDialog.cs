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


    public partial class TutorialDialog : Form
    {
        public TutorialDialog()
        {
            InitializeComponent();
            this.CenterToParent();

        }

        private void wizardPage2_Commit(object sender, AeroWizard.WizardPageConfirmEventArgs e)
        {

        }

        private void wizardControl1_Cancelling(object sender, CancelEventArgs e)
        {
            this.Close();
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
            if (kryptonCheckBox2.Checked == true)
            {
                this.Opacity = 0.8D;
            }
            else
            {
                this.Opacity = 100;
            }
        }

        private void wizardControl1_SelectedPageChanged(object sender, EventArgs e)
        {

        }

        private void TutorialDialog_Load(object sender, EventArgs e)
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

        }
    }
}
