namespace RTEditor
{
    partial class ThirdPartyDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ThirdPartyDialog));
            this.kryptonPalette1 = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.kryptonNavigator1 = new ComponentFactory.Krypton.Navigator.KryptonNavigator();
            this.kryptonPage1 = new ComponentFactory.Krypton.Navigator.KryptonPage();
            this.kryptonRichTextBox1 = new ComponentFactory.Krypton.Toolkit.KryptonRichTextBox();
            this.kryptonPage2 = new ComponentFactory.Krypton.Navigator.KryptonPage();
            this.kryptonRichTextBox2 = new ComponentFactory.Krypton.Toolkit.KryptonRichTextBox();
            this.kryptonPanel3 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonCheckButton2 = new ComponentFactory.Krypton.Toolkit.KryptonCheckButton();
            this.kryptonCheckButton1 = new ComponentFactory.Krypton.Toolkit.KryptonCheckButton();
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPanel2 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonButton1 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonNavigator1)).BeginInit();
            this.kryptonNavigator1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPage1)).BeginInit();
            this.kryptonPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPage2)).BeginInit();
            this.kryptonPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel3)).BeginInit();
            this.kryptonPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).BeginInit();
            this.kryptonPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonPalette1
            // 
            this.kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            // 
            // kryptonNavigator1
            // 
            this.kryptonNavigator1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;
            this.kryptonNavigator1.Button.CloseButtonAction = ComponentFactory.Krypton.Navigator.CloseButtonAction.None;
            this.kryptonNavigator1.Button.CloseButtonDisplay = ComponentFactory.Krypton.Navigator.ButtonDisplay.Hide;
            this.kryptonNavigator1.Location = new System.Drawing.Point(180, 12);
            this.kryptonNavigator1.Name = "kryptonNavigator1";
            this.kryptonNavigator1.Pages.AddRange(new ComponentFactory.Krypton.Navigator.KryptonPage[] {
            this.kryptonPage1,
            this.kryptonPage2});
            this.kryptonNavigator1.Palette = this.kryptonPalette1;
            this.kryptonNavigator1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonNavigator1.SelectedIndex = 0;
            this.kryptonNavigator1.Size = new System.Drawing.Size(647, 465);
            this.kryptonNavigator1.TabIndex = 1;
            this.kryptonNavigator1.Text = "kryptonNavigator1";
            // 
            // kryptonPage1
            // 
            this.kryptonPage1.AutoHiddenSlideSize = new System.Drawing.Size(200, 200);
            this.kryptonPage1.Controls.Add(this.kryptonRichTextBox1);
            this.kryptonPage1.Flags = 65534;
            this.kryptonPage1.LastVisibleSet = true;
            this.kryptonPage1.MinimumSize = new System.Drawing.Size(50, 50);
            this.kryptonPage1.Name = "kryptonPage1";
            this.kryptonPage1.Size = new System.Drawing.Size(645, 438);
            this.kryptonPage1.Text = "サードパーティコンポーネント";
            this.kryptonPage1.ToolTipTitle = "Page ToolTip";
            this.kryptonPage1.UniqueName = "A0FDF9B2B20A461428A9ED9C35E9F101";
            // 
            // kryptonRichTextBox1
            // 
            this.kryptonRichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonRichTextBox1.Location = new System.Drawing.Point(0, 0);
            this.kryptonRichTextBox1.Name = "kryptonRichTextBox1";
            this.kryptonRichTextBox1.Palette = this.kryptonPalette1;
            this.kryptonRichTextBox1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonRichTextBox1.Size = new System.Drawing.Size(645, 438);
            this.kryptonRichTextBox1.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.kryptonRichTextBox1.TabIndex = 0;
            this.kryptonRichTextBox1.Text = resources.GetString("kryptonRichTextBox1.Text");
            this.kryptonRichTextBox1.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.kryptonRichTextBox1_LinkClicked);
            // 
            // kryptonPage2
            // 
            this.kryptonPage2.AutoHiddenSlideSize = new System.Drawing.Size(200, 200);
            this.kryptonPage2.Controls.Add(this.kryptonRichTextBox2);
            this.kryptonPage2.Flags = 65534;
            this.kryptonPage2.LastVisibleSet = true;
            this.kryptonPage2.MinimumSize = new System.Drawing.Size(50, 50);
            this.kryptonPage2.Name = "kryptonPage2";
            this.kryptonPage2.Size = new System.Drawing.Size(645, 438);
            this.kryptonPage2.Text = "ライセンス";
            this.kryptonPage2.ToolTipTitle = "Page ToolTip";
            this.kryptonPage2.UniqueName = "6CFCC34C8D90415407B63651CD4D11D8";
            // 
            // kryptonRichTextBox2
            // 
            this.kryptonRichTextBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonRichTextBox2.Location = new System.Drawing.Point(0, 0);
            this.kryptonRichTextBox2.Name = "kryptonRichTextBox2";
            this.kryptonRichTextBox2.Palette = this.kryptonPalette1;
            this.kryptonRichTextBox2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonRichTextBox2.Size = new System.Drawing.Size(645, 438);
            this.kryptonRichTextBox2.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.kryptonRichTextBox2.TabIndex = 1;
            this.kryptonRichTextBox2.Text = resources.GetString("kryptonRichTextBox2.Text");
            this.kryptonRichTextBox2.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.kryptonRichTextBox2_LinkClicked);
            // 
            // kryptonPanel3
            // 
            this.kryptonPanel3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.kryptonPanel3.Controls.Add(this.kryptonCheckButton2);
            this.kryptonPanel3.Controls.Add(this.kryptonCheckButton1);
            this.kryptonPanel3.Location = new System.Drawing.Point(12, 12);
            this.kryptonPanel3.Name = "kryptonPanel3";
            this.kryptonPanel3.Palette = this.kryptonPalette1;
            this.kryptonPanel3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel3.PanelBackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.ControlClient;
            this.kryptonPanel3.Size = new System.Drawing.Size(162, 466);
            this.kryptonPanel3.TabIndex = 0;
            // 
            // kryptonCheckButton2
            // 
            this.kryptonCheckButton2.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.ListItem;
            this.kryptonCheckButton2.Dock = System.Windows.Forms.DockStyle.Top;
            this.kryptonCheckButton2.Location = new System.Drawing.Point(0, 29);
            this.kryptonCheckButton2.Name = "kryptonCheckButton2";
            this.kryptonCheckButton2.Palette = this.kryptonPalette1;
            this.kryptonCheckButton2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonCheckButton2.Size = new System.Drawing.Size(162, 29);
            this.kryptonCheckButton2.TabIndex = 2;
            this.kryptonCheckButton2.Values.Text = "ライセンス";
            this.kryptonCheckButton2.Click += new System.EventHandler(this.kryptonCheckButton2_Click);
            // 
            // kryptonCheckButton1
            // 
            this.kryptonCheckButton1.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.ListItem;
            this.kryptonCheckButton1.Checked = true;
            this.kryptonCheckButton1.Dock = System.Windows.Forms.DockStyle.Top;
            this.kryptonCheckButton1.Location = new System.Drawing.Point(0, 0);
            this.kryptonCheckButton1.Name = "kryptonCheckButton1";
            this.kryptonCheckButton1.Palette = this.kryptonPalette1;
            this.kryptonCheckButton1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonCheckButton1.Size = new System.Drawing.Size(162, 29);
            this.kryptonCheckButton1.TabIndex = 1;
            this.kryptonCheckButton1.Values.Text = "サードパーティコンポーネント";
            this.kryptonCheckButton1.Click += new System.EventHandler(this.kryptonCheckButton1_Click);
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.kryptonNavigator1);
            this.kryptonPanel1.Controls.Add(this.kryptonPanel3);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.Palette = this.kryptonPalette1;
            this.kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel1.PanelBackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.PanelRibbonInactive;
            this.kryptonPanel1.Size = new System.Drawing.Size(840, 536);
            this.kryptonPanel1.TabIndex = 2;
            // 
            // kryptonPanel2
            // 
            this.kryptonPanel2.Controls.Add(this.kryptonButton1);
            this.kryptonPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel2.Location = new System.Drawing.Point(0, 488);
            this.kryptonPanel2.Name = "kryptonPanel2";
            this.kryptonPanel2.Palette = this.kryptonPalette1;
            this.kryptonPanel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel2.Size = new System.Drawing.Size(840, 48);
            this.kryptonPanel2.TabIndex = 3;
            // 
            // kryptonButton1
            // 
            this.kryptonButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonButton1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.kryptonButton1.Location = new System.Drawing.Point(737, 11);
            this.kryptonButton1.Name = "kryptonButton1";
            this.kryptonButton1.Palette = this.kryptonPalette1;
            this.kryptonButton1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton1.Size = new System.Drawing.Size(90, 25);
            this.kryptonButton1.TabIndex = 0;
            this.kryptonButton1.Values.Text = "閉じる";
            // 
            // ThirdPartyDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(840, 536);
            this.Controls.Add(this.kryptonPanel2);
            this.Controls.Add(this.kryptonPanel1);
            this.MinimizeBox = false;
            this.Name = "ThirdPartyDialog";
            this.Palette = this.kryptonPalette1;
            this.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "サードパーティコンポーネントの一覧";
            this.Load += new System.EventHandler(this.ThirdPartyDialog_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonNavigator1)).EndInit();
            this.kryptonNavigator1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPage1)).EndInit();
            this.kryptonPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPage2)).EndInit();
            this.kryptonPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel3)).EndInit();
            this.kryptonPanel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).EndInit();
            this.kryptonPanel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private ComponentFactory.Krypton.Toolkit.KryptonPalette kryptonPalette1;
        private ComponentFactory.Krypton.Navigator.KryptonNavigator kryptonNavigator1;
        private ComponentFactory.Krypton.Navigator.KryptonPage kryptonPage1;
        private ComponentFactory.Krypton.Toolkit.KryptonRichTextBox kryptonRichTextBox1;
        private ComponentFactory.Krypton.Navigator.KryptonPage kryptonPage2;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel3;
        private ComponentFactory.Krypton.Toolkit.KryptonCheckButton kryptonCheckButton2;
        private ComponentFactory.Krypton.Toolkit.KryptonCheckButton kryptonCheckButton1;
        private ComponentFactory.Krypton.Toolkit.KryptonRichTextBox kryptonRichTextBox2;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel2;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton1;
    }
}