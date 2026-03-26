namespace RTEditor
{
    partial class InsertGraphDialog
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPalette1 = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.kryptonLabel3 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonComboBox2 = new ComponentFactory.Krypton.Toolkit.KryptonComboBox();
            this.chart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.kryptonLabel2 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonComboBox1 = new ComponentFactory.Krypton.Toolkit.KryptonComboBox();
            this.kryptonButton4 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonButton3 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonButton2 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonButton1 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonComboBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonComboBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.kryptonLabel3);
            this.kryptonPanel1.Controls.Add(this.kryptonComboBox2);
            this.kryptonPanel1.Controls.Add(this.chart);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel2);
            this.kryptonPanel1.Controls.Add(this.kryptonComboBox1);
            this.kryptonPanel1.Controls.Add(this.kryptonButton4);
            this.kryptonPanel1.Controls.Add(this.kryptonButton3);
            this.kryptonPanel1.Controls.Add(this.kryptonButton2);
            this.kryptonPanel1.Controls.Add(this.kryptonButton1);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.Palette = this.kryptonPalette1;
            this.kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel1.Size = new System.Drawing.Size(806, 450);
            this.kryptonPanel1.TabIndex = 1;
            // 
            // kryptonPalette1
            // 
            this.kryptonPalette1.AllowFormChrome = ComponentFactory.Krypton.Toolkit.InheritBool.True;
            this.kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            // 
            // kryptonLabel3
            // 
            this.kryptonLabel3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonLabel3.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel3.Location = new System.Drawing.Point(656, 96);
            this.kryptonLabel3.Name = "kryptonLabel3";
            this.kryptonLabel3.Palette = this.kryptonPalette1;
            this.kryptonLabel3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel3.Size = new System.Drawing.Size(75, 20);
            this.kryptonLabel3.TabIndex = 10;
            this.kryptonLabel3.Values.Text = "表のスタイル:";
            // 
            // kryptonComboBox2
            // 
            this.kryptonComboBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.kryptonComboBox2.DropDownWidth = 141;
            this.kryptonComboBox2.Items.AddRange(new object[] {
            "散布図（点プロット）",
            "高速描画用の散布図",
            "バブルチャート（点の大きさで値を表現）",
            "折れ線グラフ",
            "高速描画用折れ線グラフ",
            "スプライン曲線",
            "階段状の折れ線グラフ",
            "面グラフ",
            "スプライン曲線の面グラフ",
            "積み上げ面グラフ",
            "積み上げ面グラフ（割合 100%）",
            "横棒グラフ",
            "積み上げ横棒グラフ",
            "積み上げ横棒グラフ（割合 100%）",
            "縦棒グラフ",
            "積み上げ縦棒グラフ",
            "積み上げ縦棒グラフ（割合 100%）",
            "円グラフ",
            "ドーナツグラフ",
            "株価チャート（OHLC）",
            "ローソク足チャート",
            "範囲チャート（上下の値を帯で表現）",
            "スプライン曲線の範囲チャート",
            "横方向の範囲チャート",
            "縦方向の範囲チャート",
            "レーダーチャート",
            "極座標チャート",
            "誤差範囲付きチャート",
            "箱ひげ図",
            "ピラミッドチャート",
            "ファネルチャート",
            "三本足チャート",
            "陰陽足チャート",
            "イント＆フィギュア",
            "レンコチャート"});
            this.kryptonComboBox2.Location = new System.Drawing.Point(656, 122);
            this.kryptonComboBox2.Name = "kryptonComboBox2";
            this.kryptonComboBox2.Palette = this.kryptonPalette1;
            this.kryptonComboBox2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonComboBox2.Size = new System.Drawing.Size(141, 21);
            this.kryptonComboBox2.TabIndex = 9;
            this.kryptonComboBox2.Text = "縦棒グラフ";
            this.kryptonComboBox2.SelectedIndexChanged += new System.EventHandler(this.kryptonComboBox2_SelectedIndexChanged);
            // 
            // chart
            // 
            chartArea1.Name = "ChartArea1";
            this.chart.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chart.Legends.Add(legend1);
            this.chart.Location = new System.Drawing.Point(12, 12);
            this.chart.Name = "chart";
            this.chart.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None;
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart.Series.Add(series1);
            this.chart.Size = new System.Drawing.Size(638, 426);
            this.chart.TabIndex = 8;
            this.chart.Text = "chart1";
            // 
            // kryptonLabel2
            // 
            this.kryptonLabel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonLabel2.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel2.Location = new System.Drawing.Point(656, 43);
            this.kryptonLabel2.Name = "kryptonLabel2";
            this.kryptonLabel2.Palette = this.kryptonPalette1;
            this.kryptonLabel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel2.Size = new System.Drawing.Size(50, 20);
            this.kryptonLabel2.TabIndex = 7;
            this.kryptonLabel2.Values.Text = "表の色:";
            // 
            // kryptonComboBox1
            // 
            this.kryptonComboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.kryptonComboBox1.DropDownWidth = 141;
            this.kryptonComboBox1.Items.AddRange(new object[] {
            "なし",
            "Bright",
            "Grayscale\t",
            "Excel",
            "Light",
            "Pastel",
            "EarthTones",
            "SemiTransparent",
            "Berry",
            "Chocolate",
            "Fire",
            "SeaGreen\t",
            "BrightPastel"});
            this.kryptonComboBox1.Location = new System.Drawing.Point(656, 69);
            this.kryptonComboBox1.Name = "kryptonComboBox1";
            this.kryptonComboBox1.Palette = this.kryptonPalette1;
            this.kryptonComboBox1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonComboBox1.Size = new System.Drawing.Size(141, 21);
            this.kryptonComboBox1.TabIndex = 6;
            this.kryptonComboBox1.Text = "なし";
            this.kryptonComboBox1.SelectedIndexChanged += new System.EventHandler(this.kryptonComboBox1_SelectedIndexChanged);
            // 
            // kryptonButton4
            // 
            this.kryptonButton4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonButton4.Location = new System.Drawing.Point(707, 413);
            this.kryptonButton4.Name = "kryptonButton4";
            this.kryptonButton4.Palette = this.kryptonPalette1;
            this.kryptonButton4.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton4.Size = new System.Drawing.Size(90, 25);
            this.kryptonButton4.TabIndex = 4;
            this.kryptonButton4.Values.Text = "ヘルプ";
            this.kryptonButton4.Click += new System.EventHandler(this.kryptonButton4_Click);
            // 
            // kryptonButton3
            // 
            this.kryptonButton3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonButton3.Location = new System.Drawing.Point(707, 382);
            this.kryptonButton3.Name = "kryptonButton3";
            this.kryptonButton3.Palette = this.kryptonPalette1;
            this.kryptonButton3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton3.Size = new System.Drawing.Size(90, 25);
            this.kryptonButton3.TabIndex = 3;
            this.kryptonButton3.Values.Text = "キャンセル";
            this.kryptonButton3.Click += new System.EventHandler(this.kryptonButton3_Click);
            // 
            // kryptonButton2
            // 
            this.kryptonButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonButton2.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.kryptonButton2.Location = new System.Drawing.Point(707, 351);
            this.kryptonButton2.Name = "kryptonButton2";
            this.kryptonButton2.Palette = this.kryptonPalette1;
            this.kryptonButton2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton2.Size = new System.Drawing.Size(90, 25);
            this.kryptonButton2.TabIndex = 2;
            this.kryptonButton2.Values.Text = "適用";
            this.kryptonButton2.Click += new System.EventHandler(this.kryptonButton2_Click);
            // 
            // kryptonButton1
            // 
            this.kryptonButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonButton1.Location = new System.Drawing.Point(656, 12);
            this.kryptonButton1.Name = "kryptonButton1";
            this.kryptonButton1.Palette = this.kryptonPalette1;
            this.kryptonButton1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton1.Size = new System.Drawing.Size(141, 25);
            this.kryptonButton1.TabIndex = 1;
            this.kryptonButton1.Values.Text = "CSVファイル読み込み...";
            this.kryptonButton1.Click += new System.EventHandler(this.kryptonButton1_Click);
            // 
            // InsertGraphDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(806, 450);
            this.Controls.Add(this.kryptonPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InsertGraphDialog";
            this.Palette = this.kryptonPalette1;
            this.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.StateCommon.Header.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StateCommon.Header.Content.ShortText.Hint = ComponentFactory.Krypton.Toolkit.PaletteTextHint.ClearTypeGridFit;
            this.Text = "グラフの挿入";
            this.Load += new System.EventHandler(this.InsertGraphDialog_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            this.kryptonPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonComboBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonComboBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton4;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton3;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton2;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel2;
        private ComponentFactory.Krypton.Toolkit.KryptonComboBox kryptonComboBox1;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel3;
        private ComponentFactory.Krypton.Toolkit.KryptonComboBox kryptonComboBox2;
        private ComponentFactory.Krypton.Toolkit.KryptonPalette kryptonPalette1;
    }
}