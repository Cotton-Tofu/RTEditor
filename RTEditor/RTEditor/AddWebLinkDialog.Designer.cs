namespace RTEditor
{
    partial class AddWebLinkDialog
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddWebLinkDialog));
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonTextBox2 = new ComponentFactory.Krypton.Toolkit.KryptonTextBox();
            this.buttonSpecAny4 = new ComponentFactory.Krypton.Toolkit.ButtonSpecAny();
            this.buttonSpecAny2 = new ComponentFactory.Krypton.Toolkit.ButtonSpecAny();
            this.buttonSpecAny3 = new ComponentFactory.Krypton.Toolkit.ButtonSpecAny();
            this.kryptonPalette1 = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.kryptonLabel6 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonLabel5 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonButton1 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonLabel3 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonTextBox1 = new ComponentFactory.Krypton.Toolkit.KryptonTextBox();
            this.buttonSpecAny1 = new ComponentFactory.Krypton.Toolkit.ButtonSpecAny();
            this.kryptonLabel2 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonLabel1 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.kryptonPanel2 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonCheckBox1 = new ComponentFactory.Krypton.Toolkit.KryptonCheckBox();
            this.kryptonLabel4 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonButton3 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonButton2 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).BeginInit();
            this.kryptonPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.kryptonTextBox2);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel6);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel5);
            this.kryptonPanel1.Controls.Add(this.kryptonButton1);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel3);
            this.kryptonPanel1.Controls.Add(this.kryptonTextBox1);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel2);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel1);
            this.kryptonPanel1.Controls.Add(this.webView21);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.Palette = this.kryptonPalette1;
            this.kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel1.PanelBackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.PanelRibbonInactive;
            this.kryptonPanel1.Size = new System.Drawing.Size(902, 635);
            this.kryptonPanel1.TabIndex = 0;
            // 
            // kryptonTextBox2
            // 
            this.kryptonTextBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonTextBox2.ButtonSpecs.AddRange(new ComponentFactory.Krypton.Toolkit.ButtonSpecAny[] {
            this.buttonSpecAny4,
            this.buttonSpecAny2,
            this.buttonSpecAny3});
            this.kryptonTextBox2.Location = new System.Drawing.Point(276, 38);
            this.kryptonTextBox2.Name = "kryptonTextBox2";
            this.kryptonTextBox2.Palette = this.kryptonPalette1;
            this.kryptonTextBox2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonTextBox2.ReadOnly = true;
            this.kryptonTextBox2.Size = new System.Drawing.Size(614, 28);
            this.kryptonTextBox2.StateCommon.Content.Font = new System.Drawing.Font("Yu Gothic UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.kryptonTextBox2.TabIndex = 10;
            // 
            // buttonSpecAny4
            // 
            this.buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
            this.buttonSpecAny4.Image = ((System.Drawing.Image)(resources.GetObject("buttonSpecAny4.Image")));
            this.buttonSpecAny4.ToolTipBody = "既定のブラウザーで表示";
            this.buttonSpecAny4.UniqueName = "B192E2E853BF4F0A1EBC5AEDD232695B";
            this.buttonSpecAny4.Click += new System.EventHandler(this.buttonSpecAny4_Click);
            // 
            // buttonSpecAny2
            // 
            this.buttonSpecAny2.Image = ((System.Drawing.Image)(resources.GetObject("buttonSpecAny2.Image")));
            this.buttonSpecAny2.ToolTipBody = "ページの再読み込み";
            this.buttonSpecAny2.UniqueName = "03CC135737C847FBF7A94F99CECD5C3A";
            this.buttonSpecAny2.Click += new System.EventHandler(this.kryptonButton1_Click);
            // 
            // buttonSpecAny3
            // 
            this.buttonSpecAny3.Image = ((System.Drawing.Image)(resources.GetObject("buttonSpecAny3.Image")));
            this.buttonSpecAny3.ToolTipBody = "ページを閉じる";
            this.buttonSpecAny3.UniqueName = "4A306F2C4D14422733A5361C7F3CB913";
            this.buttonSpecAny3.Click += new System.EventHandler(this.buttonSpecAny1_Click);
            // 
            // kryptonPalette1
            // 
            this.kryptonPalette1.AllowFormChrome = ComponentFactory.Krypton.Toolkit.InheritBool.True;
            this.kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.ProfessionalSystem;
            // 
            // kryptonLabel6
            // 
            this.kryptonLabel6.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel6.Location = new System.Drawing.Point(12, 192);
            this.kryptonLabel6.Name = "kryptonLabel6";
            this.kryptonLabel6.Palette = this.kryptonPalette1;
            this.kryptonLabel6.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel6.Size = new System.Drawing.Size(154, 22);
            this.kryptonLabel6.StateCommon.ShortText.Color1 = System.Drawing.Color.Red;
            this.kryptonLabel6.TabIndex = 8;
            this.kryptonLabel6.Values.Text = "URLの形式が正しくありません";
            this.kryptonLabel6.Visible = false;
            // 
            // kryptonLabel5
            // 
            this.kryptonLabel5.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel5.Location = new System.Drawing.Point(276, 14);
            this.kryptonLabel5.Name = "kryptonLabel5";
            this.kryptonLabel5.Palette = this.kryptonPalette1;
            this.kryptonLabel5.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel5.Size = new System.Drawing.Size(55, 22);
            this.kryptonLabel5.TabIndex = 7;
            this.kryptonLabel5.Values.Text = "プレビュー:";
            // 
            // kryptonButton1
            // 
            this.kryptonButton1.Location = new System.Drawing.Point(12, 157);
            this.kryptonButton1.Name = "kryptonButton1";
            this.kryptonButton1.Palette = this.kryptonPalette1;
            this.kryptonButton1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton1.Size = new System.Drawing.Size(121, 25);
            this.kryptonButton1.TabIndex = 6;
            this.kryptonButton1.Values.Text = "表示";
            this.kryptonButton1.Click += new System.EventHandler(this.kryptonButton1_Click);
            // 
            // kryptonLabel3
            // 
            this.kryptonLabel3.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel3.Location = new System.Drawing.Point(12, 95);
            this.kryptonLabel3.Name = "kryptonLabel3";
            this.kryptonLabel3.Palette = this.kryptonPalette1;
            this.kryptonLabel3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel3.Size = new System.Drawing.Size(35, 22);
            this.kryptonLabel3.TabIndex = 4;
            this.kryptonLabel3.Values.Text = "URL:";
            // 
            // kryptonTextBox1
            // 
            this.kryptonTextBox1.ButtonSpecs.AddRange(new ComponentFactory.Krypton.Toolkit.ButtonSpecAny[] {
            this.buttonSpecAny1});
            this.kryptonTextBox1.Location = new System.Drawing.Point(12, 121);
            this.kryptonTextBox1.Name = "kryptonTextBox1";
            this.kryptonTextBox1.Palette = this.kryptonPalette1;
            this.kryptonTextBox1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonTextBox1.Size = new System.Drawing.Size(255, 25);
            this.kryptonTextBox1.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.kryptonTextBox1.TabIndex = 3;
            this.kryptonTextBox1.TabStop = false;
            this.kryptonTextBox1.Text = "https://";
            this.kryptonTextBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.kryptonTextBox1_KeyDown);
            // 
            // buttonSpecAny1
            // 
            this.buttonSpecAny1.Type = ComponentFactory.Krypton.Toolkit.PaletteButtonSpecStyle.Close;
            this.buttonSpecAny1.UniqueName = "01906E01EFCC460B2CAB9B87EB709718";
            this.buttonSpecAny1.Click += new System.EventHandler(this.buttonSpecAny1_Click);
            // 
            // kryptonLabel2
            // 
            this.kryptonLabel2.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel2.Location = new System.Drawing.Point(12, 55);
            this.kryptonLabel2.Name = "kryptonLabel2";
            this.kryptonLabel2.Palette = this.kryptonPalette1;
            this.kryptonLabel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel2.Size = new System.Drawing.Size(172, 22);
            this.kryptonLabel2.TabIndex = 2;
            this.kryptonLabel2.Values.Text = "使用するURLを入力してください。";
            // 
            // kryptonLabel1
            // 
            this.kryptonLabel1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.TitlePanel;
            this.kryptonLabel1.Location = new System.Drawing.Point(12, 20);
            this.kryptonLabel1.Name = "kryptonLabel1";
            this.kryptonLabel1.Palette = this.kryptonPalette1;
            this.kryptonLabel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel1.Size = new System.Drawing.Size(159, 32);
            this.kryptonLabel1.StateCommon.ShortText.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.kryptonLabel1.StateCommon.ShortText.Hint = ComponentFactory.Krypton.Toolkit.PaletteTextHint.ClearTypeGridFit;
            this.kryptonLabel1.TabIndex = 1;
            this.kryptonLabel1.Values.Text = "Webリンクを追加";
            // 
            // webView21
            // 
            this.webView21.AllowExternalDrop = true;
            this.webView21.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.webView21.CreationProperties = null;
            this.webView21.DefaultBackgroundColor = System.Drawing.Color.White;
            this.webView21.Location = new System.Drawing.Point(276, 72);
            this.webView21.Name = "webView21";
            this.webView21.Size = new System.Drawing.Size(614, 503);
            this.webView21.Source = new System.Uri("about:blank", System.UriKind.Absolute);
            this.webView21.TabIndex = 0;
            this.webView21.ZoomFactor = 1D;
            this.webView21.NavigationStarting += new System.EventHandler<Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs>(this.webView21_NavigationStarting);
            this.webView21.NavigationCompleted += new System.EventHandler<Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs>(this.webView21_NavigationCompleted);
            this.webView21.ContentLoading += new System.EventHandler<Microsoft.Web.WebView2.Core.CoreWebView2ContentLoadingEventArgs>(this.webView21_ContentLoading);
            // 
            // kryptonPanel2
            // 
            this.kryptonPanel2.Controls.Add(this.kryptonCheckBox1);
            this.kryptonPanel2.Controls.Add(this.kryptonLabel4);
            this.kryptonPanel2.Controls.Add(this.kryptonButton3);
            this.kryptonPanel2.Controls.Add(this.kryptonButton2);
            this.kryptonPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel2.Location = new System.Drawing.Point(0, 585);
            this.kryptonPanel2.Name = "kryptonPanel2";
            this.kryptonPanel2.Palette = this.kryptonPalette1;
            this.kryptonPanel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel2.Size = new System.Drawing.Size(902, 50);
            this.kryptonPanel2.TabIndex = 0;
            // 
            // kryptonCheckBox1
            // 
            this.kryptonCheckBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonCheckBox1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonCheckBox1.Location = new System.Drawing.Point(518, 14);
            this.kryptonCheckBox1.Name = "kryptonCheckBox1";
            this.kryptonCheckBox1.Palette = this.kryptonPalette1;
            this.kryptonCheckBox1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonCheckBox1.Size = new System.Drawing.Size(134, 22);
            this.kryptonCheckBox1.TabIndex = 4;
            this.kryptonCheckBox1.Values.Text = "正確なURLを使用する";
            // 
            // kryptonLabel4
            // 
            this.kryptonLabel4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.kryptonLabel4.Location = new System.Drawing.Point(12, 13);
            this.kryptonLabel4.Name = "kryptonLabel4";
            this.kryptonLabel4.Palette = this.kryptonPalette1;
            this.kryptonLabel4.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel4.Size = new System.Drawing.Size(319, 22);
            this.kryptonLabel4.TabIndex = 3;
            this.kryptonLabel4.Values.Text = "ヒント:リッチテキストエディタに直接URLを入力することもできます。";
            // 
            // kryptonButton3
            // 
            this.kryptonButton3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonButton3.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.kryptonButton3.Location = new System.Drawing.Point(662, 12);
            this.kryptonButton3.Name = "kryptonButton3";
            this.kryptonButton3.Palette = this.kryptonPalette1;
            this.kryptonButton3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton3.Size = new System.Drawing.Size(111, 25);
            this.kryptonButton3.TabIndex = 1;
            this.kryptonButton3.Values.Text = "使用して入力";
            this.kryptonButton3.Click += new System.EventHandler(this.kryptonButton3_Click);
            // 
            // kryptonButton2
            // 
            this.kryptonButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonButton2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.kryptonButton2.Location = new System.Drawing.Point(779, 12);
            this.kryptonButton2.Name = "kryptonButton2";
            this.kryptonButton2.Palette = this.kryptonPalette1;
            this.kryptonButton2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton2.Size = new System.Drawing.Size(111, 25);
            this.kryptonButton2.TabIndex = 0;
            this.kryptonButton2.Values.Text = "キャンセル";
            // 
            // AddWebLinkDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(902, 635);
            this.Controls.Add(this.kryptonPanel2);
            this.Controls.Add(this.kryptonPanel1);
            this.MinimizeBox = false;
            this.Name = "AddWebLinkDialog";
            this.Palette = this.kryptonPalette1;
            this.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.StateCommon.Header.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StateCommon.Header.Content.ShortText.Hint = ComponentFactory.Krypton.Toolkit.PaletteTextHint.ClearTypeGridFit;
            this.Text = "Webリンク挿入";
            this.Load += new System.EventHandler(this.AddWebLinkDialog_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            this.kryptonPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).EndInit();
            this.kryptonPanel2.ResumeLayout(false);
            this.kryptonPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel3;
        private ComponentFactory.Krypton.Toolkit.KryptonTextBox kryptonTextBox1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel2;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel1;
        private Microsoft.Web.WebView2.WinForms.WebView2 webView21;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel2;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton1;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton2;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton3;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel4;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel5;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel6;
        private ComponentFactory.Krypton.Toolkit.ButtonSpecAny buttonSpecAny1;
        private ComponentFactory.Krypton.Toolkit.KryptonTextBox kryptonTextBox2;
        private ComponentFactory.Krypton.Toolkit.ButtonSpecAny buttonSpecAny4;
        private ComponentFactory.Krypton.Toolkit.ButtonSpecAny buttonSpecAny2;
        private ComponentFactory.Krypton.Toolkit.ButtonSpecAny buttonSpecAny3;
        private ComponentFactory.Krypton.Toolkit.KryptonPalette kryptonPalette1;
        private ComponentFactory.Krypton.Toolkit.KryptonCheckBox kryptonCheckBox1;
    }
}