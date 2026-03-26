namespace RTEditor
{
    partial class AboutBox1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutBox1));
            this.tableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.logoPictureBox = new System.Windows.Forms.PictureBox();
            this.labelProductName = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonPalette1 = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.labelVersion = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.labelCopyright = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.labelCompanyName = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.okButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonPanel2 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPanel3 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonToolkitPoweredByControl1 = new Krypton.Toolkit.KryptonToolkitPoweredByControl();
            this.kryptonLabel1 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonGroupBox1 = new ComponentFactory.Krypton.Toolkit.KryptonGroupBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.kryptonLabel2 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonLinkLabel1 = new ComponentFactory.Krypton.Toolkit.KryptonLinkLabel();
            this.tableLayoutPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).BeginInit();
            this.kryptonPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonGroupBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonGroupBox1.Panel)).BeginInit();
            this.kryptonGroupBox1.Panel.SuspendLayout();
            this.kryptonGroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel
            // 
            this.tableLayoutPanel.BackColor = System.Drawing.Color.Transparent;
            this.tableLayoutPanel.ColumnCount = 2;
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33F));
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 67F));
            this.tableLayoutPanel.Controls.Add(this.logoPictureBox, 0, 0);
            this.tableLayoutPanel.Controls.Add(this.labelProductName, 1, 0);
            this.tableLayoutPanel.Controls.Add(this.labelVersion, 1, 1);
            this.tableLayoutPanel.Controls.Add(this.labelCopyright, 1, 2);
            this.tableLayoutPanel.Controls.Add(this.labelCompanyName, 1, 3);
            this.tableLayoutPanel.Controls.Add(this.okButton, 1, 5);
            this.tableLayoutPanel.Controls.Add(this.kryptonPanel2, 1, 4);
            this.tableLayoutPanel.Location = new System.Drawing.Point(7, 8);
            this.tableLayoutPanel.Name = "tableLayoutPanel";
            this.tableLayoutPanel.RowCount = 6;
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel.Size = new System.Drawing.Size(456, 283);
            this.tableLayoutPanel.TabIndex = 0;
            // 
            // logoPictureBox
            // 
            this.logoPictureBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("logoPictureBox.Image")));
            this.logoPictureBox.Location = new System.Drawing.Point(3, 3);
            this.logoPictureBox.Name = "logoPictureBox";
            this.tableLayoutPanel.SetRowSpan(this.logoPictureBox, 6);
            this.logoPictureBox.Size = new System.Drawing.Size(144, 277);
            this.logoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.logoPictureBox.TabIndex = 12;
            this.logoPictureBox.TabStop = false;
            // 
            // labelProductName
            // 
            this.labelProductName.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            this.labelProductName.Location = new System.Drawing.Point(153, 3);
            this.labelProductName.Name = "labelProductName";
            this.labelProductName.Palette = this.kryptonPalette1;
            this.labelProductName.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.labelProductName.Size = new System.Drawing.Size(49, 20);
            this.labelProductName.TabIndex = 25;
            this.labelProductName.Values.Text = "製品名";
            // 
            // kryptonPalette1
            // 
            this.kryptonPalette1.AllowFormChrome = ComponentFactory.Krypton.Toolkit.InheritBool.True;
            this.kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            // 
            // labelVersion
            // 
            this.labelVersion.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.labelVersion.Location = new System.Drawing.Point(153, 31);
            this.labelVersion.Name = "labelVersion";
            this.labelVersion.Palette = this.kryptonPalette1;
            this.labelVersion.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.labelVersion.Size = new System.Drawing.Size(59, 20);
            this.labelVersion.TabIndex = 26;
            this.labelVersion.Values.Text = "バージョン";
            // 
            // labelCopyright
            // 
            this.labelCopyright.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.labelCopyright.Location = new System.Drawing.Point(153, 59);
            this.labelCopyright.Name = "labelCopyright";
            this.labelCopyright.Palette = this.kryptonPalette1;
            this.labelCopyright.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.labelCopyright.Size = new System.Drawing.Size(49, 20);
            this.labelCopyright.TabIndex = 27;
            this.labelCopyright.Values.Text = "著作権";
            // 
            // labelCompanyName
            // 
            this.labelCompanyName.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.labelCompanyName.Location = new System.Drawing.Point(153, 87);
            this.labelCompanyName.Name = "labelCompanyName";
            this.labelCompanyName.Palette = this.kryptonPalette1;
            this.labelCompanyName.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.labelCompanyName.Size = new System.Drawing.Size(49, 20);
            this.labelCompanyName.TabIndex = 28;
            this.labelCompanyName.Values.Text = "会社名";
            // 
            // okButton
            // 
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.okButton.Dock = System.Windows.Forms.DockStyle.Right;
            this.okButton.Location = new System.Drawing.Point(363, 256);
            this.okButton.Name = "okButton";
            this.okButton.Palette = this.kryptonPalette1;
            this.okButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.okButton.Size = new System.Drawing.Size(90, 24);
            this.okButton.TabIndex = 29;
            this.okButton.Values.Text = "OK";
            // 
            // kryptonPanel2
            // 
            this.kryptonPanel2.AutoScroll = true;
            this.kryptonPanel2.Controls.Add(this.kryptonGroupBox1);
            this.kryptonPanel2.Controls.Add(this.kryptonPanel3);
            this.kryptonPanel2.Controls.Add(this.kryptonToolkitPoweredByControl1);
            this.kryptonPanel2.Controls.Add(this.kryptonLabel1);
            this.kryptonPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel2.Location = new System.Drawing.Point(153, 115);
            this.kryptonPanel2.Name = "kryptonPanel2";
            this.kryptonPanel2.Palette = this.kryptonPalette1;
            this.kryptonPanel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel2.Size = new System.Drawing.Size(300, 135);
            this.kryptonPanel2.TabIndex = 30;
            // 
            // kryptonPanel3
            // 
            this.kryptonPanel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel3.Location = new System.Drawing.Point(0, 353);
            this.kryptonPanel3.Name = "kryptonPanel3";
            this.kryptonPanel3.Palette = this.kryptonPalette1;
            this.kryptonPanel3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel3.Size = new System.Drawing.Size(283, 21);
            this.kryptonPanel3.TabIndex = 31;
            // 
            // kryptonToolkitPoweredByControl1
            // 
            this.kryptonToolkitPoweredByControl1.Location = new System.Drawing.Point(14, 56);
            this.kryptonToolkitPoweredByControl1.Name = "kryptonToolkitPoweredByControl1";
            this.kryptonToolkitPoweredByControl1.Size = new System.Drawing.Size(258, 161);
            this.kryptonToolkitPoweredByControl1.TabIndex = 30;
            this.kryptonToolkitPoweredByControl1.ToolkitType = Krypton.Toolkit.ToolkitType.Stable;
            // 
            // kryptonLabel1
            // 
            this.kryptonLabel1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel1.Location = new System.Drawing.Point(13, 13);
            this.kryptonLabel1.Name = "kryptonLabel1";
            this.kryptonLabel1.Palette = this.kryptonPalette1;
            this.kryptonLabel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel1.Size = new System.Drawing.Size(254, 36);
            this.kryptonLabel1.TabIndex = 29;
            this.kryptonLabel1.Values.Text = "このソフトウェアはC#のWindows Formsで作成し\r\nオープンソースとしてGitHubで公開されています。";
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.tableLayoutPanel);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.Palette = this.kryptonPalette1;
            this.kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel1.Size = new System.Drawing.Size(470, 299);
            this.kryptonPanel1.TabIndex = 1;
            // 
            // kryptonGroupBox1
            // 
            this.kryptonGroupBox1.Location = new System.Drawing.Point(12, 228);
            this.kryptonGroupBox1.Name = "kryptonGroupBox1";
            this.kryptonGroupBox1.Palette = this.kryptonPalette1;
            this.kryptonGroupBox1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            // 
            // kryptonGroupBox1.Panel
            // 
            this.kryptonGroupBox1.Panel.Controls.Add(this.kryptonLinkLabel1);
            this.kryptonGroupBox1.Panel.Controls.Add(this.kryptonLabel2);
            this.kryptonGroupBox1.Panel.Controls.Add(this.pictureBox1);
            this.kryptonGroupBox1.Size = new System.Drawing.Size(260, 125);
            this.kryptonGroupBox1.TabIndex = 32;
            this.kryptonGroupBox1.Values.Heading = "アイコン権利表記";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(15, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(50, 50);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // kryptonLabel2
            // 
            this.kryptonLabel2.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel2.Location = new System.Drawing.Point(76, 8);
            this.kryptonLabel2.Name = "kryptonLabel2";
            this.kryptonLabel2.Palette = this.kryptonPalette1;
            this.kryptonLabel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel2.Size = new System.Drawing.Size(165, 52);
            this.kryptonLabel2.TabIndex = 29;
            this.kryptonLabel2.Values.Text = "他のソフトウェアで使用している\r\nアイコンはIcons 8社のPichon\r\nを使用しています。";
            // 
            // kryptonLinkLabel1
            // 
            this.kryptonLinkLabel1.Location = new System.Drawing.Point(80, 65);
            this.kryptonLinkLabel1.Name = "kryptonLinkLabel1";
            this.kryptonLinkLabel1.Palette = this.kryptonPalette1;
            this.kryptonLinkLabel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLinkLabel1.Size = new System.Drawing.Size(72, 20);
            this.kryptonLinkLabel1.TabIndex = 30;
            this.kryptonLinkLabel1.Values.Text = "Icons8.com";
            this.kryptonLinkLabel1.LinkClicked += new System.EventHandler(this.kryptonLinkLabel1_LinkClicked);
            // 
            // AboutBox1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 299);
            this.Controls.Add(this.kryptonPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutBox1";
            this.Palette = this.kryptonPalette1;
            this.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.StateCommon.Header.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StateCommon.Header.Content.ShortText.Hint = ComponentFactory.Krypton.Toolkit.PaletteTextHint.ClearTypeGridFit;
            this.Text = "AboutBox1";
            this.Load += new System.EventHandler(this.AboutBox1_Load);
            this.tableLayoutPanel.ResumeLayout(false);
            this.tableLayoutPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.logoPictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).EndInit();
            this.kryptonPanel2.ResumeLayout(false);
            this.kryptonPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonGroupBox1.Panel)).EndInit();
            this.kryptonGroupBox1.Panel.ResumeLayout(false);
            this.kryptonGroupBox1.Panel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonGroupBox1)).EndInit();
            this.kryptonGroupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel;
        private System.Windows.Forms.PictureBox logoPictureBox;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel labelProductName;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel labelVersion;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel labelCopyright;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel labelCompanyName;
        private ComponentFactory.Krypton.Toolkit.KryptonButton okButton;
        private ComponentFactory.Krypton.Toolkit.KryptonPalette kryptonPalette1;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel2;
        private Krypton.Toolkit.KryptonToolkitPoweredByControl kryptonToolkitPoweredByControl1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel1;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel3;
        private ComponentFactory.Krypton.Toolkit.KryptonGroupBox kryptonGroupBox1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private ComponentFactory.Krypton.Toolkit.KryptonLinkLabel kryptonLinkLabel1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel2;
    }
}
