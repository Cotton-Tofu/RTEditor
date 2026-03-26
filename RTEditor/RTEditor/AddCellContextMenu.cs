using ComponentFactory.Krypton.Toolkit;
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
    public partial class AddCellContextMenu : KryptonForm
    {
        public AddCellContextMenu()
        {
            InitializeComponent();
            
        }

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void AddCellContextMenu_Load(object sender, EventArgs e)
        {
            this.Location = Cursor.Position;
        }

        private void AddCellContextMenu_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }

        //1行目
        private void button1_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(1行×1列)";
            button1.BackColor = Color.NavajoWhite;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(1行×2列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(1行×3列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(1行×4列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
        }

        private void button5_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(1行×5列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
        }

        private void button5_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
        }

        private void button6_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(1行×6列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;
        }

        private void button6_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
        }

        private void button7_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(1行×7列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;
            button7.BackColor = Color.NavajoWhite;
        }

        private void button7_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;
        }

        //2行目
        private void button8_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(2行×1列)";
            button1.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
        }

        private void button8_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;

            button8.BackColor = Color.White;

        }

        private void button9_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(2行×2列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
        }

        private void button9_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
        }

        private void button10_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(2行×3列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
        }

        private void button10_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
        }

        private void button11_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(2行×4列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
        }

        private void button11_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
        }

        private void button12_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(2行×5列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
        }

        private void button12_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
        }

        private void button13_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(2行×6列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
            button13.BackColor = Color.NavajoWhite;
        }

        private void button13_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
        }

        private void button14_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(2行×7列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;
            button7.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
            button13.BackColor = Color.NavajoWhite;
            button14.BackColor = Color.NavajoWhite;
        }

        private void button14_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;
        }



        //3行目

        private void button15_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(3行×1列)";
            button1.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
        }

        private void button15_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
        }

        private void button16_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(3行×2列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
        }

        private void button16_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
        }

        private void button17_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(3行×3列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
        }

        private void button17_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
        }

        private void button18_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(3行×4列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
        }

        private void button18_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
        }

        private void button19_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表3行×5列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
            button19.BackColor = Color.NavajoWhite;
        }

        private void button19_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
        }

        private void button20_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(3行×6列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
            button13.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
            button19.BackColor = Color.NavajoWhite;
            button20.BackColor = Color.NavajoWhite;
        }

        private void button20_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
        }
        private void button21_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(3行×7列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;
            button7.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
            button13.BackColor = Color.NavajoWhite;
            button14.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
            button19.BackColor = Color.NavajoWhite;
            button20.BackColor = Color.NavajoWhite;
            button21.BackColor = Color.NavajoWhite;
        }

        private void button21_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;
        }

        //4行目

        private void button22_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(4行×1列)";
            button1.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;

            button22.BackColor = Color.NavajoWhite;
        }

        private void button22_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
        }

        private void button23_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(4行×2列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;

            button22.BackColor= Color.NavajoWhite;
            button23.BackColor = Color.NavajoWhite;
        }

        private void button23_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
        }

        private void button24_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(4行×3列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;

            button22.BackColor = Color.NavajoWhite;
            button23.BackColor = Color.NavajoWhite;
            button24.BackColor = Color.NavajoWhite;
        }

        private void button24_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
        }

        private void button25_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(4行×4列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;

            button22.BackColor = Color.NavajoWhite;
            button23.BackColor = Color.NavajoWhite;
            button24.BackColor = Color.NavajoWhite;
            button25.BackColor = Color.NavajoWhite;
        }

        private void button25_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
        }

        private void button26_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(4行×5列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
            button19.BackColor = Color.NavajoWhite;

            button22.BackColor = Color.NavajoWhite;
            button23.BackColor = Color.NavajoWhite;
            button24.BackColor = Color.NavajoWhite;
            button25.BackColor = Color.NavajoWhite;
            button26.BackColor = Color.NavajoWhite;
        }

        private void button26_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
        }

        private void button27_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(4行×6列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
            button13.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
            button19.BackColor = Color.NavajoWhite;
            button20.BackColor = Color.NavajoWhite;

            button22.BackColor = Color.NavajoWhite;
            button23.BackColor = Color.NavajoWhite;
            button24.BackColor = Color.NavajoWhite;
            button25.BackColor = Color.NavajoWhite;
            button26.BackColor = Color.NavajoWhite;
            button27.BackColor = Color.NavajoWhite;
        }

        private void button27_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
        }

        private void button28_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(4行×7列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;
            button7.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
            button13.BackColor = Color.NavajoWhite;
            button14.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
            button19.BackColor = Color.NavajoWhite;
            button20.BackColor = Color.NavajoWhite;
            button21.BackColor = Color.NavajoWhite;

            button22.BackColor = Color.NavajoWhite;
            button23.BackColor = Color.NavajoWhite;
            button24.BackColor = Color.NavajoWhite;
            button25.BackColor = Color.NavajoWhite;
            button26.BackColor = Color.NavajoWhite;
            button27.BackColor = Color.NavajoWhite;
            button28.BackColor = Color.NavajoWhite;
        }

        private void button28_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
        }

        private void button29_MouseEnter(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表(5行×1列)";
            button1.BackColor = Color.NavajoWhite;
            button2.BackColor = Color.NavajoWhite;
            button3.BackColor = Color.NavajoWhite;
            button4.BackColor = Color.NavajoWhite;
            button5.BackColor = Color.NavajoWhite;
            button6.BackColor = Color.NavajoWhite;

            button8.BackColor = Color.NavajoWhite;
            button9.BackColor = Color.NavajoWhite;
            button10.BackColor = Color.NavajoWhite;
            button11.BackColor = Color.NavajoWhite;
            button12.BackColor = Color.NavajoWhite;
            button13.BackColor = Color.NavajoWhite;

            button15.BackColor = Color.NavajoWhite;
            button16.BackColor = Color.NavajoWhite;
            button17.BackColor = Color.NavajoWhite;
            button18.BackColor = Color.NavajoWhite;
            button19.BackColor = Color.NavajoWhite;
            button20.BackColor = Color.NavajoWhite;

            button22.BackColor = Color.NavajoWhite;
            button23.BackColor = Color.NavajoWhite;
            button24.BackColor = Color.NavajoWhite;
            button25.BackColor = Color.NavajoWhite;
            button26.BackColor = Color.NavajoWhite;
            button27.BackColor = Color.NavajoWhite;
            button28.BackColor = Color.NavajoWhite;

            button29.BackColor = Color.NavajoWhite;

        }

        private void button29_MouseLeave(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "表の挿入";
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;

            button8.BackColor = Color.White;
            button9.BackColor = Color.White;
            button10.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;

            button15.BackColor = Color.White;
            button16.BackColor = Color.White;
            button17.BackColor = Color.White;
            button18.BackColor = Color.White;
            button19.BackColor = Color.White;
            button20.BackColor = Color.White;
            button21.BackColor = Color.White;

            button22.BackColor = Color.White;
            button23.BackColor = Color.White;
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;

            button29.BackColor = Color.White;
            button30.BackColor = Color.White;
        }


        //5行目

    }
}
