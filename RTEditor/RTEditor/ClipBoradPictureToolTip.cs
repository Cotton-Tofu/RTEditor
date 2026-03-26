using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RTEditor
{
    public partial class ClipBoradPictureToolTip : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public ClipBoradPictureToolTip()
        {
            InitializeComponent();
        }
        public Image image { get; set; }

        private void ClipBoradPictureToolTip_Load(object sender, EventArgs e)
        {
            pictureBox1.Image = image;
        }

        private void ClipBoradPictureToolTip_Activated(object sender, EventArgs e)
        {
            pictureBox1.Image = image;
        }
    }
}
