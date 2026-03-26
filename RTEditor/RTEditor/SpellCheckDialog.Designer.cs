namespace RTEditor
{
    partial class SpellCheckDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPanel2 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonLabel1 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonRichTextBox1 = new ComponentFactory.Krypton.Toolkit.KryptonRichTextBox();
            this.kryptonListBox1 = new ComponentFactory.Krypton.Toolkit.KryptonListBox();
            this.kryptonLabel2 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonLabel3 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonListBox2 = new ComponentFactory.Krypton.Toolkit.KryptonListBox();
            this.kryptonButton1 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonButton2 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.spellCheckerAdv1 = new Syncfusion.Windows.Forms.Tools.SpellCheckerAdv();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).BeginInit();
            this.SuspendLayout();
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.kryptonButton2);
            this.kryptonPanel1.Controls.Add(this.kryptonButton1);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel3);
            this.kryptonPanel1.Controls.Add(this.kryptonListBox2);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel2);
            this.kryptonPanel1.Controls.Add(this.kryptonListBox1);
            this.kryptonPanel1.Controls.Add(this.kryptonRichTextBox1);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel1);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.PanelBackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.PanelRibbonInactive;
            this.kryptonPanel1.Size = new System.Drawing.Size(602, 253);
            this.kryptonPanel1.TabIndex = 0;
            // 
            // kryptonPanel2
            // 
            this.kryptonPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel2.Location = new System.Drawing.Point(0, 253);
            this.kryptonPanel2.Name = "kryptonPanel2";
            this.kryptonPanel2.Size = new System.Drawing.Size(602, 44);
            this.kryptonPanel2.TabIndex = 1;
            // 
            // kryptonLabel1
            // 
            this.kryptonLabel1.Location = new System.Drawing.Point(12, 16);
            this.kryptonLabel1.Name = "kryptonLabel1";
            this.kryptonLabel1.Size = new System.Drawing.Size(115, 20);
            this.kryptonLabel1.TabIndex = 0;
            this.kryptonLabel1.Values.Text = "選択箇所のテキスト:";
            // 
            // kryptonRichTextBox1
            // 
            this.kryptonRichTextBox1.Location = new System.Drawing.Point(12, 42);
            this.kryptonRichTextBox1.Name = "kryptonRichTextBox1";
            this.kryptonRichTextBox1.Size = new System.Drawing.Size(219, 196);
            this.kryptonRichTextBox1.TabIndex = 1;
            this.kryptonRichTextBox1.Text = "kryptonRichTextBox1";
            // 
            // kryptonListBox1
            // 
            this.kryptonListBox1.Location = new System.Drawing.Point(250, 42);
            this.kryptonListBox1.Name = "kryptonListBox1";
            this.kryptonListBox1.Size = new System.Drawing.Size(217, 82);
            this.kryptonListBox1.TabIndex = 2;
            // 
            // kryptonLabel2
            // 
            this.kryptonLabel2.Location = new System.Drawing.Point(250, 16);
            this.kryptonLabel2.Name = "kryptonLabel2";
            this.kryptonLabel2.Size = new System.Drawing.Size(112, 20);
            this.kryptonLabel2.TabIndex = 3;
            this.kryptonLabel2.Values.Text = "修正が必要な候補:";
            // 
            // kryptonLabel3
            // 
            this.kryptonLabel3.Location = new System.Drawing.Point(250, 130);
            this.kryptonLabel3.Name = "kryptonLabel3";
            this.kryptonLabel3.Size = new System.Drawing.Size(112, 20);
            this.kryptonLabel3.TabIndex = 5;
            this.kryptonLabel3.Values.Text = "修正が必要な候補:";
            // 
            // kryptonListBox2
            // 
            this.kryptonListBox2.Location = new System.Drawing.Point(250, 156);
            this.kryptonListBox2.Name = "kryptonListBox2";
            this.kryptonListBox2.Size = new System.Drawing.Size(217, 82);
            this.kryptonListBox2.TabIndex = 4;
            // 
            // kryptonButton1
            // 
            this.kryptonButton1.Location = new System.Drawing.Point(473, 42);
            this.kryptonButton1.Name = "kryptonButton1";
            this.kryptonButton1.Size = new System.Drawing.Size(120, 25);
            this.kryptonButton1.TabIndex = 6;
            this.kryptonButton1.Values.Text = "修正候補を置換";
            // 
            // kryptonButton2
            // 
            this.kryptonButton2.Location = new System.Drawing.Point(473, 73);
            this.kryptonButton2.Name = "kryptonButton2";
            this.kryptonButton2.Size = new System.Drawing.Size(120, 25);
            this.kryptonButton2.TabIndex = 7;
            this.kryptonButton2.Values.Text = "すべて置換";
            // 
            // spellCheckerAdv1
            // 
            this.spellCheckerAdv1.Culture = null;
            this.spellCheckerAdv1.CustomDictionaryPath = "C:\\ProgramData\\Microsoft Corporation\\Microsoft® Visual Studio®\\18.2.11408.102\\Cus" +
    "tom_Dictionay.dic";
            this.spellCheckerAdv1.Dictionaries = null;
            this.spellCheckerAdv1.DictionaryPath = "Syncfusion_en_us.dic";
            // 
            // SpellCheckDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 297);
            this.Controls.Add(this.kryptonPanel1);
            this.Controls.Add(this.kryptonPanel2);
            this.Name = "SpellCheckDialog";
            this.Text = "スペルチェッカー";
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            this.kryptonPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel2;
        private ComponentFactory.Krypton.Toolkit.KryptonListBox kryptonListBox1;
        private ComponentFactory.Krypton.Toolkit.KryptonRichTextBox kryptonRichTextBox1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel1;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel2;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton2;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel3;
        private ComponentFactory.Krypton.Toolkit.KryptonListBox kryptonListBox2;
        private Syncfusion.Windows.Forms.Tools.SpellCheckerAdv spellCheckerAdv1;
    }
}