using ComponentFactory.Krypton.Docking;
using ComponentFactory.Krypton.Navigator;
using ComponentFactory.Krypton.Ribbon;
using ComponentFactory.Krypton.Toolkit;
using ComponentFactory.Krypton.Workspace;
using Google.GenAI;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Taskbar;
using Ookii.Dialogs.WinForms;
using OpenAI.Chat;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Tools.Enums;
using Syncfusion.Windows.Forms.Tools.Renderers;
using Syncfusion.WinForms.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Media;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Windows.Shell;
using Windows.Security.Credentials.UI;
using static Microsoft.WindowsAPICodePack.ApplicationServices.ApplicationRestartRecoveryManager;
using static Microsoft.WindowsAPICodePack.ApplicationServices.AppRestartRecoveryNativeMethods;
using static RTEditor.Form1;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;


namespace RTEditor
{
    public partial class Form1 : ComponentFactory.Krypton.Toolkit.KryptonForm
    {


        //リボンタブ以外の機能の各プログラム


        //編集用のMdiフォームとリッチテキストボックスを作成
        public KryptonForm form;
        public RichTextBox richTextBox;

        public Form1()
        {
            //UIカルチャーを日本語にローカライズする
            System.Threading.Thread.CurrentThread.CurrentUICulture = new CultureInfo("ja-JP");
            InitializeComponent();
        }



        // --- 既存の DllImport セクションの近くに追加 ---
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int SetCurrentProcessExplicitAppUserModelID(string appID);

        private void Form1_Shown(object sender, EventArgs e)
        {
            // 1) プロセスに AppUserModelID を設定（"com.yourcompany.rteditor" は実際の識別子に置き換える）
            try
            {
                SetCurrentProcessExplicitAppUserModelID("com.cottontofu.rteditor");
            }
            catch
            {
                // 失敗しても続行可能（管理者権限などで失敗するケースがある）
            }

            // 2) JumpList を作成して「最近使ったもの」を表示する
            jumpList = Microsoft.WindowsAPICodePack.Taskbar.JumpList.CreateJumpList();
            jumpList.KnownCategoryToDisplay = JumpListKnownCategoryType.Recent;
            jumpList.Refresh();

            if(this.MdiChildren.Length == 0)
            {
                DisabledRibbonAllGroup(sender, e);
            }
            else
            {
                EnabledRibbonAllGroup(sender, e);
            }

            if (Properties.Settings.Default.IsUseRichTextBoxDarkMode == true)
            {
                miniColorPicker2.AutomaticColor = Color.Black;
            }
            else if (Properties.Settings.Default.IsUseRichTextBoxDarkMode == false)
            {
                miniColorPicker2.AutomaticColor = Color.White;
            }
        }


        private void Form1_Activated(object sender, EventArgs e)
        {
            clipBoradStatus.Start();

        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            clipBoradStatus.Stop();

            if (this.WindowState == FormWindowState.Minimized)
            {
                if (fm2 != null)
                {
                    fm2.WindowState = FormWindowState.Minimized;
                }

            }
        }



        //初めてこのプログラムが起動したか確かめるためのbool型のフラグ
        public bool FarstExecute { get; set; } = true;

        //新しいドキュメントを作成するためのメソッド
        public void MdiFormCreate(object sender, EventArgs e)
        {

            //Mdiフォームとテキストボックスのインスタンスを作成

            form = new KryptonForm();

            form.Palette = kryptonPalette2;
            form.MdiParent = this;
            form.AllowFormChrome = true;
            form.StateCommon.Header.Content.ShortText.Font = new Font("Yu Gothic UI", 9);
            form.StateCommon.Header.Content.ShortText.Hint = PaletteTextHint.ClearTypeGridFit;
            form.StateCommon.Header.Content.LongText.Font = new Font("Yu Gothic UI", 9);
            form.StateCommon.Header.Content.LongText.Hint = PaletteTextHint.ClearTypeGridFit;

            //初めて起動したときのみMdiウィンドウを最大化する
            if (FarstExecute == true)
            {
                form.WindowState = FormWindowState.Maximized;
                FarstExecute = false;
            }
            else
            {
                //以前開いたMdiフォームに応じて新しいMdiウィンドウを最大化・復元するか決める
                if (form.WindowState == FormWindowState.Maximized)
                {
                    form.WindowState = FormWindowState.Maximized;
                }
                else if (form.WindowState == FormWindowState.Normal)
                {
                    form.WindowState = FormWindowState.Normal;
                }
                else if (form.WindowState == FormWindowState.Minimized)
                {
                    form.WindowState = FormWindowState.Minimized;
                }
            }

            form.FormBorderStyle = FormBorderStyle.Sizable;
            form.ShowIcon = false;

            //リッチテキストボックス
            richTextBox = new RichTextBox();
            richTextBox.Dock = DockStyle.Fill;
            richTextBox.BorderStyle = BorderStyle.None;
            richTextBox.ScrollBars = RichTextBoxScrollBars.ForcedBoth;
            richTextBox.AcceptsTab = true;
            richTextBox.DetectUrls = true;



            //イベントハンドラの追加
            form.FormClosing += form_FormClosing;
            form.FormClosed += Form_FormClosed;
            form.Activated += form_Activated;
            richTextBox.TextChanged += richTextBox_TextChanged;
            richTextBox.SelectionChanged += richTextBox_SelectionChanged;
            richTextBox.AllowDrop = true;
            richTextBox.DragDrop += richTextBox_DragDrop;
            richTextBox.DragEnter += Form1_DragEnter;
            richTextBox.DragLeave += Form1_DragLeave;
            richTextBox.LinkClicked += richTextBox_LinkClicked;
            richTextBox.MouseDown += RichTextBox_MouseDown;
            richTextBox.MouseUp += richTextBox_MouseUp;

            contextMenuStrip1.Closed += richTextBox_MinitoolBarCloseAction;
            richTextBox.MouseClick += richTextBox_MinitoolBarCloseAction;
            richTextBox.Click += richTextBox_MinitoolBarCloseAction;
            richTextBox.KeyDown += richTextBox_MinitoolBarCloseAction;
            richTextBox.DragEnter += richTextBox_MinitoolBarCloseAction;


            //既定のフォントとスタイル

            //既定のフォント名
            richTextBox.Font = new Font(Properties.Settings.Default.DefaultFontName, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style);

            //既定のフォントサイズ
            richTextBox.Font = new Font(richTextBox.SelectionFont.Name, Properties.Settings.Default.DefaultFontSize, richTextBox.SelectionFont.Style);

            //太字
            if (Properties.Settings.Default.DefaultFontIsBold == true)
            {
                richTextBox.Font = new Font(richTextBox.SelectionFont.Name, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style | FontStyle.Bold);
            }

            //斜体
            if (Properties.Settings.Default.DefaultFontIsItalic == true)
            {
                richTextBox.Font = new Font(richTextBox.SelectionFont.Name, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style | FontStyle.Italic);
            }

            //下線
            if (Properties.Settings.Default.DefaultFontIsUnderline == true)
            {
                richTextBox.Font = new Font(richTextBox.SelectionFont.Name, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style | FontStyle.Underline);
            }

            //打ち消し線
            if (Properties.Settings.Default.DefaultFontIsStrikeout == true)
            {
                richTextBox.Font = new Font(richTextBox.SelectionFont.Name, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style | FontStyle.Strikeout);
            }



            //表示タブ

            //ドキュメントの表示オプション

            //既定でドキュメント内の文字列の折り返し
            if (Properties.Settings.Default.IsUseWordWarp == true)
            {
                richTextBox.WordWrap = true;
            }
            else if (Properties.Settings.Default.IsUseWordWarp == false)
            {
                richTextBox.WordWrap = false;
            }

            //単語自動選択機能
            if (Properties.Settings.Default.IsUseAutoWordSelection == true)
            {
                richTextBox.AutoWordSelection = true;
            }
            else if (Properties.Settings.Default.IsUseAutoWordSelection == false)
            {
                richTextBox.AutoWordSelection = false;
            }

            if (Properties.Settings.Default.IsUseRichTextBoxDarkMode == true)
            {
                richTextBox.BackColor = System.Drawing.Color.Black;
            }
            else if (Properties.Settings.Default.IsUseRichTextBoxDarkMode == false)
            {
                richTextBox.BackColor = System.Drawing.Color.White;
            }


            //リッチテキストボックスの言語オプションを設定
            richTextBox.LanguageOption &= ~RichTextBoxLanguageOptions.DualFont;

            //Mdiフォームとリッチテキストエディタのプロパティを設定して表示
            form.Text = "無題";

            //リッチテキストボックスをformに追加
            form.Controls.Add(richTextBox);

            richTextBox.ContextMenuStrip = contextMenuStrip1;

            //formを表示する
            form.Show();

            //左端の段落選択用の空白の表示
            if (Properties.Settings.Default.IsShowSelectionMargin == true)
            {
                richTextBox.ShowSelectionMargin = true;
            }
            else if (Properties.Settings.Default.IsShowSelectionMargin == false)
            {
                richTextBox.ShowSelectionMargin = false;
            }

            richTextBox_ScanFontStyle();

            kryptonTrackBar1.Value = 5;

            //文字色
            richTextBox.ForeColor = Properties.Settings.Default.DefaultFontTextColor;
            //蛍光ペン
            richTextBox.SelectionBackColor = Properties.Settings.Default.DefaultFontTextHilightColor;

            if(this.MdiChildren.Length > 0)
            {
                EnabledRibbonAllGroup(sender, e);
            }

        }

        //MDIを完全に閉じたときにMDIの数が1であれば(全部閉じた場合)リボンのUIの状態をfalseにする
        private void Form_FormClosed(object sender, EventArgs e)
        {
            if(this.MdiChildren.Length == 1)
            {
                DisabledRibbonAllGroup(sender, e);
            }
        }

        private void richTextBox_MinitoolBarCloseAction(object sender, EventArgs e)
        {
            IsUsingMiniToolBar = false;
            miniToolBar1.Hide();
            radialMenu1.MenuVisibility = false;
            radialMenu1.HidePopup();
        }
        public bool isMouseUp { get; set; } = true;
        //リッチテキストボックスでマウスドラッグしたあとにミニツールバーを表示する処理
        public void richTextBox_MouseUp(object sender, MouseEventArgs e)
        {
            if (isMouseUp == true)
            {
                Form activeform = this.ActiveMdiChild;
                richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                if (richTextBox.SelectionLength > 0)
                {
                    var screen = richTextBox.PointToScreen(new Point(e.X, e.Y + richTextBox.Cursor.Size.Height));
                    if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "MiniToolBar")
                    {
                        miniToolBar1.Show(screen);
                    }
                    else if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "RadialMenu")
                    {
                        try
                        {
                            radialMenu1.MenuVisibility = false;
                            radialMenu1.ShowRadialMenu(screen);
                        }
                        catch
                        {//何もしない
                        }

                    }


                    if (e.Button == MouseButtons.Right && richTextBox.SelectionLength > 0)
                    {
                        ShowWordLikePopup(e.Location);
                    }
                }
                else
                {
                    miniToolBar1.Hide();
                    radialMenu1.MenuVisibility = false;
                    radialMenu1.HidePopup();
                }
            }

        }

        private void ShowWordLikePopup(Point clientPoint)
        {
            Point screen = richTextBox.PointToScreen(clientPoint);

            miniToolBar1.Hide();
            radialMenu1.HidePopup();

            // ContextMenuをその下に
            Point menuPos = new Point(
                screen.X,
                screen.Y
            );

            contextMenuStrip1.Show(menuPos);
        }


        private void RichTextBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (isMouseUp == false)
            {
                if (e.Button == MouseButtons.Right)
                {
                    miniToolBar1.Hide();
                    radialMenu1.MenuVisibility = false;
                    radialMenu1.HidePopup();
                }
            }

        }

        public bool IsUsingMiniToolBar { get; set; }
        //ミニツールバーが閉じられたときにコンテキストメニューも閉じる処理
        private void miniToolBar1_Closing(object sender, ToolStripDropDownClosingEventArgs e)
        {
            if (IsUsingMiniToolBar == true)
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
                contextMenuStrip1.Close();
                IsMiniToolBarVisible = false;
            }

        }

        //リッチテキストボックス内のURLをクリックしたときの処理
        public void richTextBox_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            Process.Start(e.LinkText);
        }

        //Form1にファイルがドロップされたときの処理
        private void richTextBox_DragDrop(object sender, DragEventArgs e)
        {
            int fileCount = 0;
            DataObject data = new DataObject();
            while (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                Form activeform = this.ActiveMdiChild;
                richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                string filepath1 = files[fileCount];
                data.SetData(DataFormats.FileDrop, new string[] { filepath1 });
                Clipboard.SetDataObject(data, true);

                richTextBox.Paste();
                fileCount++;
                if (fileCount == files.Length)
                {
                    fileCount = 0;
                    break;
                }

            }
        }

        private void kryptonRibbonGroupButton2_Click(object sender, EventArgs e)
        {

        }

        //Mdiウィンドウを閉じたときの処理
        //親ウィンドウを閉じたときにすべての子ウィンドウを閉じる（子ウィンドウ単体ではすべて閉じない）
        public bool IsFormClosing { get; set; } = false;

        public int count { get; set; } = 0;
        Form activeform { get; set; }
        String fname { get; set; }
        DialogResult result;
        Microsoft.WindowsAPICodePack.Dialogs.TaskDialog SavetaskDialog;
        public void form_FormClosing(object sender1, FormClosingEventArgs e)
        {

            //Mdiウィンドウが存在する場合かつ変更されている内容がある場合、保存確認ダイアログを表示する
            if (this.MdiChildren.Length > 0)
            {
                try
                {
                    if (IsFormClosing == true)
                    {
                        activeform = this.MdiChildren[count];
                        richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                        fname = activeform.Text;
                    }
                    else if (IsFormClosing == false)
                    {
                        activeform = (Form)sender1;
                        richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                        fname = activeform.Text;
                    }

                    if (activeform.Text.Contains("*") == true)
                    {
                        activeform.Activate();
                        string str = activeform.Text.Replace("*", "");
                        SavetaskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog();

                        SavetaskDialog.Caption = "保存確認";
                        SavetaskDialog.InstructionText = str + "の内容は変更されています。変更内容を保存しますか？";
                        if (File.Exists(str) == false)
                        {
                            SavetaskDialog.Text = "保存する場合、このドキュメントに名前を付けて保存します。";
                        }
                        else
                        {
                            SavetaskDialog.Text = "保存する場合、このドキュメントを上書き保存します。";
                        }
                        SavetaskDialog.StandardButtons = TaskDialogStandardButtons.Yes | TaskDialogStandardButtons.No | TaskDialogStandardButtons.Cancel;

                        SavetaskDialog.Icon = TaskDialogStandardIcon.None;
                        SavetaskDialog.OwnerWindowHandle = this.Handle;
                        TaskDialogResult tdResult = SavetaskDialog.Show();

                        if (tdResult == TaskDialogResult.Yes)
                        {
                            result = DialogResult.Yes;
                            // 既存のファイルパスが存在する場合はそのまま上書き保存しカウンターをインクリメントして閉じる処理を続行
                            if (File.Exists(str))
                            {
                                if (str.Contains(".rtf"))
                                {
                                    richTextBox?.SaveFile(str, RichTextBoxStreamType.RichText);
                                }
                                else if (str.Contains(".txt"))
                                {
                                    richTextBox?.SaveFile(str, RichTextBoxStreamType.PlainText);
                                }
                                activeform.Text = str;
                                count++;
                            }
                            // 既存のファイルパスが存在しない場合は保存ダイアログを表示して名前を付けて保存
                            else
                            {
                                // 保存処理
                                using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Title = "保存する場所を選択...", Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt", FilterIndex = Properties.Settings.Default.DefaultSaveMode + 1 })
                                {
                                    string str1 = activeform.Text.Replace("*", "");
                                    if (File.Exists(str1) == true)
                                    {
                                        saveFileDialog.FileName = str1;
                                    }
                                    else
                                    {
                                        if (richTextBox?.Text.Length > 0)
                                        {
                                            saveFileDialog.FileName = richTextBox?.Lines[0];
                                        }
                                        else
                                        {
                                            saveFileDialog.FileName = "無題";
                                        }
                                    }

                                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                    {
                                        try
                                        {
                                            if (saveFileDialog.FilterIndex == 1)
                                            {
                                                richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.RichText);
                                            }
                                            else if (saveFileDialog.FilterIndex == 2)
                                            {
                                                richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
                                            }

                                            //最近使用したドキュメントのファイルパス保存・表示処理
                                            if (Properties.Settings.Default.RecentDocsPath.Contains(saveFileDialog.FileName) == false)
                                            {
                                                AddRecentDocs(saveFileDialog.FileName);
                                                SaveRecentDocs(saveFileDialog.FileName);
                                            }

                                            // 保存が完了したらタイトルから*を外す
                                            activeform.Text = saveFileDialog.FileName;
                                            // Mdiウィンドウが1つだけの場合、そのウィンドウを閉じ破棄する
                                            if (this.MdiChildren.Length == 0)
                                            {
                                                activeform.Hide();
                                                activeform.Dispose();
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("ファイルの保存中にエラーが発生しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            // 保存ダイアログがキャンセルされた場合、閉じる処理を中止するが保存確認は続行する
                                            e.Cancel = true;
                                        }
                                    }
                                    else
                                    {
                                        // 保存ダイアログがキャンセルされた場合、閉じる処理を中止するが保存確認は続行する
                                        e.Cancel = true;
                                    }
                                    //カウントをインクリメントする
                                    count++;

                                }
                            }
                        }
                        else if (tdResult == TaskDialogResult.No)
                        {
                            result = DialogResult.No;
                            // いいえが選択された場合、保存せずカウンターをインクリメントして閉じる処理を続行
                            count++;
                            // すべてのドキュメントを閉じるときにキャンセルされなかった場合、親ウィンドウも閉じる
                            if (IsFormClosing == true && count == this.MdiChildren.Length)
                            {
                                this.Close();
                            }
                        }
                        else if (tdResult == TaskDialogResult.Cancel)
                        {
                            result = DialogResult.Cancel;
                            // カウントをリセット
                            count = 0;
                            // キャンセルされた場合、閉じる処理を中止
                            e.Cancel = true;
                            IsFormClosing = false;
                        }


                    }
                }
                catch
                {
                    e.Cancel = true;
                }

            }
            else
            {
                SavetaskDialog.Dispose();
            }

        }

        //親ウィンドウを閉じたときの保存確認後の処理
        private void Form1_FormClosing(object sender, FormClosingEventArgs a)
        {
            IsFormClosing = true;

            if (IsFormClosing == true)
            {
                if (this.MdiChildren.Length > 0)
                {
                    if (activeform.Text.Contains("*") == true)
                    {
                        if (result == DialogResult.Yes)
                        {
                            form_FormClosing(sender, a);
                            a.Cancel = false;
                            IsFormClosing = true;
                            activeform.Activate();
                        }
                        else if (result == DialogResult.No)
                        {
                            form_FormClosing(sender, a);
                            a.Cancel = false;
                            IsFormClosing = true;
                            activeform.Activate();
                        }
                        else if (result == DialogResult.Cancel)
                        {
                            a.Cancel = true;
                            IsFormClosing = false;
                            count = 0;
                        }

                    }
                    else
                    {
                        a.Cancel = false;
                        count = 0;
                    }
                }
                //すべてのMdiウィンドウを閉じた後、クリップボードの内容を監視するタイマーを停止する
                else
                {
                    clipBoradStatus.Stop();
                }

            }
        }

        //リッチテキストボックスのテキストが変更されたらそのMdiウィンドウを閉じたときに保存確認ダイアログを表示する
        private void richTextBox_TextChanged(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            //ファイルを読み込んでいる場合、タイトルをファイルパスにする
            if (File.Exists(activeform.Text.Replace("*", "")) == true)
            {
                //タイトルをファイルパスにする
                if (richTextBox.Text != ReadOnlyRTB.Text)
                {
                    activeform.Text = "*" + filepath;
                }
                else
                {
                    activeform.Text = filepath;
                }
            }
            //ファイルを読み込んでいない場合、タイトルを無題にするかリッチテキストボックスの1行目のテキストにする
            else
            {

                //タイトルをリッチテキストボックスの1行目にする
                if (richTextBox.Text != string.Empty)
                {
                    string str = richTextBox.Lines[0];

                    if (str.Length < 20)
                    {
                        activeform.Text = "*" + str;

                    }
                    else
                    {
                        string s = str.Substring(0, 20);
                        activeform.Text = "*" + s + "...";
                    }
                }
                else
                {
                    activeform.Text = "無題";
                }

            }

            //文字数と行数をステータスバーに表示
            kryptonLabel3.Text = richTextBox.Text.Length.ToString() + "文字";
            kryptonLabel4.Text = richTextBox.Lines.Length.ToString() + "行";

            //段落のインデントと間隔
            kryptonRibbonGroupNumericUpDown1.Value = richTextBox.SelectionIndent;
            kryptonRibbonGroupNumericUpDown2.Value = richTextBox.SelectionRightIndent;
            kryptonRibbonGroupNumericUpDown3.Value = richTextBox.SelectionCharOffset;
        }

        //他のMdiウィンドウに変更したときに逐次そのファイルの内容とリッチテキストボックス編集状態とReadOnlyRTBの内容を照合する。それに応じて保存するかしないか決める
        private void form_Activated(object sender, EventArgs e)
        {
            if (IsFileOpening == false)
            {
                Form activeform = this.ActiveMdiChild;
                string filename = activeform.Text.Replace("*", "");
                if (File.Exists(filename))
                {
                    ReadOnlyRTB.Clear();

                    if (filename.Contains(".rtf") | filename.Contains(".txt.rtf"))
                    {
                        ReadOnlyRTB.LoadFile(filename, RichTextBoxStreamType.RichText);
                    }
                    else if (filename.Contains(".txt") | filename.Contains(".rtf.txt"))
                    {
                        ReadOnlyRTB.LoadFile(filename, RichTextBoxStreamType.PlainText);
                    }

                    filepath = filename;
                    richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    //文字数と行数をステータスバーに表示
                   
                    if (richTextBox.Text != ReadOnlyRTB.Text)
                    {
                        activeform.Text = "*" + filepath;
                    }
                    else
                    {
                        activeform.Text = filepath;
                    }
                }
                richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                kryptonLabel3.Text = richTextBox.Text.Length.ToString() + "文字";
                kryptonLabel4.Text = richTextBox.Lines.Length.ToString() + "行";

            }

            //ズーム率を現在のMdiウィンドウのリッチテキストボックスに合わせる
            activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();


            //リッチテキストのフォントスタイルの状態をスキャン
            richTextBox_ScanFontStyle();

            //検索インデックスをリセット
            index = 0;
            _lastSearchIndex = 0;


            //TTSの停止処理
            speechSynthesizer.Pause();
            speechSynthesizer.SpeakAsyncCancelAll();
        }

        //ファイルの読み込みとMdiウィンドウの表示
        public string filepath { get; set; }

        //予めテキスト比較用のリッチテキストボックスを用意しておく(表示はしない)
        public RichTextBox ReadOnlyRTB = new RichTextBox();

        public bool IsFileOpening { get; set; } = false;



        //最近使用したドキュメントの表示処理
        public void ScanRecentDocs()
        {
            System.IO.StringReader stringReader = new StringReader(Properties.Settings.Default.RecentDocsPath);
            while (stringReader.Peek() > -1)
            {
                string filepath = stringReader.ReadLine();
                if (File.Exists(filepath) == true)
                {
                    AddRecentDocs(filepath);
                }
                else
                {
                    Properties.Settings.Default.RecentDocsPath = Properties.Settings.Default.RecentDocsPath.Replace("\r\n" + filepath, "");
                    Properties.Settings.Default.Save();
                }
            }
            stringReader.Close();
            stringReader.Dispose();
        }

        //最近使用したドキュメントの候補を追加する処理
        KryptonRibbonRecentDoc[] kryptonRibbonRecentDoc;
        public void AddRecentDocs(string RecentDocFilePath)
        {
            this.kryptonRibbonRecentDoc = new KryptonRibbonRecentDoc[1];
            for (int i = 0; i < this.kryptonRibbonRecentDoc.Length; i++)
            {
                this.kryptonRibbonRecentDoc[i] = new KryptonRibbonRecentDoc();
                this.kryptonRibbonRecentDoc[i].Text = RecentDocFilePath;
                this.kryptonRibbonRecentDoc[i].Click += new EventHandler(kryptonRibbonRecentDoc_Click);
            }
            kryptonRibbon1.RibbonAppButton.AppButtonRecentDocs.AddRange(this.kryptonRibbonRecentDoc);
        }

        //最近使用したドキュメントのクリック&読み込み処理
        private void kryptonRibbonRecentDoc_Click(object sender, EventArgs e)
        {
            string filePath = ((KryptonRibbonRecentDoc)sender).Text;
            if (File.Exists(filePath) == true)
            {
                MdiFormCreate(sender, e);

                Form activeform = this.ActiveMdiChild;
                RichTextBox richTextBox = (RichTextBox)activeform.Controls[0];
                //ファイルをリッチテキストボックスに表示
                if (filePath.Contains(".rtf") | filePath.Contains(".txt.rtf"))
                {
                    richTextBox.LoadFile(filePath, RichTextBoxStreamType.RichText);
                    ReadOnlyRTB.LoadFile(filePath, RichTextBoxStreamType.RichText);
                }
                else if (filePath.Contains(".txt") | filePath.Contains(".rtf.txt"))
                {
                    richTextBox.LoadFile(filePath, RichTextBoxStreamType.PlainText);
                    ReadOnlyRTB.LoadFile(filePath, RichTextBoxStreamType.PlainText);
                }

                activeform.Text = filePath;
                filepath = filePath;
            }
            else
            {
                ScanRecentDocs();
            }

        }

        //最近使用したドキュメントのファイルパスを保存する処理
        public void SaveRecentDocs(string RecentDocFilePath)
        {
            if (Properties.Settings.Default.RecentDocsPath == string.Empty)
            {
                Properties.Settings.Default.RecentDocsPath = "\r\n" + RecentDocFilePath;
            }
            else
            {
                Properties.Settings.Default.RecentDocsPath = Properties.Settings.Default.RecentDocsPath + Environment.NewLine + RecentDocFilePath;
            }
            Properties.Settings.Default.Save();
        }

        private void kryptonContextMenuItem2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog { Title = "開くファイルを選択...", Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt|リッチテキストファイルと書式なしテキストファイル(*.rtf;*.txt)|*.rtf;*.txt", Multiselect = true, ReadOnlyChecked = true })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    IsFileOpening = true;
                    //選択したファイルの数だけMdiフォームとリッチテキストボックスを作成して表示
                    int fileCount = 0;
                    string[] selectedFiles = openFileDialog.FileNames;
                    while (selectedFiles.Length > 0)
                    {
                        MdiFormCreate(sender, e);
                        Form activeform = this.ActiveMdiChild;
                        RichTextBox richTextBox = (RichTextBox)activeform.Controls[0];

                        filepath = openFileDialog.FileName;

                        if (selectedFiles[fileCount].Contains(".rtf") | selectedFiles[fileCount].Contains(".txt.rtf"))
                        {
                            richTextBox.LoadFile(selectedFiles[fileCount], RichTextBoxStreamType.RichText);
                            ReadOnlyRTB.LoadFile(selectedFiles[fileCount], RichTextBoxStreamType.RichText);
                        }
                        else if (selectedFiles[fileCount].Contains(".txt") | selectedFiles[fileCount].Contains(".rtf.txt"))
                        {
                            richTextBox.LoadFile(selectedFiles[fileCount], RichTextBoxStreamType.PlainText);
                            ReadOnlyRTB.LoadFile(selectedFiles[fileCount], RichTextBoxStreamType.PlainText);
                        }
                        //ファイルをリッチテキストボックスに表示

                        activeform.Text = selectedFiles[fileCount];

                        //最近使用したドキュメントのファイルパス保存・表示処理
                        if (Properties.Settings.Default.RecentDocsPath.Contains(selectedFiles[fileCount]) == false)
                        {
                            AddRecentDocs(selectedFiles[fileCount]);
                            SaveRecentDocs(selectedFiles[fileCount]);
                        }

                        AddRecentFileJumplist(selectedFiles[fileCount], jumpList);

                        fileCount++;
                        //選択したファイルの数だけ繰り返す
                        if (fileCount == selectedFiles.Length)
                        {
                            break;
                        }

                    }
                    IsFileOpening = false;

                }
            }
        }

        public static void AddRecentFileJumplist(string fullFilePath, Microsoft.WindowsAPICodePack.Taskbar.JumpList jumpListName)
        {
            Microsoft.WindowsAPICodePack.Taskbar.JumpList.AddToRecent(fullFilePath);
            jumpListName.Refresh();
        }

        Microsoft.WindowsAPICodePack.Taskbar.JumpList jumpList;

        //Form1を外部から操作するためのインスタンス取得処理
        private static Form1 form1_Instace;
        public static Form1 _Form1_Instance
        {
            get { return form1_Instace; }
            set { form1_Instace = value; }
        }


        System.Windows.Forms.Timer clipBoradStatus = new System.Windows.Forms.Timer();
        private void Form1_Load(object sender, EventArgs e)
        {
            
            //クリップボード監視タイマーの実行
            clipBoradStatus.Interval = 100; // 100ミリ秒ごとにチェック
            clipBoradStatus.Tick += ClipBoradStatus_Tick;
            clipBoradStatus.Start();
            AddClipboardFormatListener(Handle);

            //最近使用したドキュメントのファイルをチェック
            ScanRecentDocs();

            Form1._Form1_Instance = this;

            InstalledFontCollection fonts = new InstalledFontCollection();
            FontFamily[] fontFamilies = fonts.Families;

            statusStripEx1.VisualStyle = Syncfusion.Windows.Forms.Tools.StatusStripExStyle.Default;

            //コンボボックスフォントの適用
            foreach (FontFamily font in fontFamilies)
            {
                kryptonRibbonGroupComboBox1.Items.Add(font.Name);
                kryptonRibbonGroupComboBox1.AutoCompleteCustomSource.Add(font.Name);
                toolStripComboBox2.Items.Add(font.Name);
                toolStripComboBox2.AutoCompleteCustomSource.Add(font.Name);
            }

            //デフォルトのリボンタブをホームタブに設定
            kryptonRibbon1.SelectedTab = kryptonRibbonTab1;

            //起動時の状態
            if (Properties.Settings.Default.AppStartupTaskMode == 1)
            {
                MdiFormCreate(sender, e);
            }

            AddColorPickerToMiniToolBar();
            //設定の復元
            AppLoadSettings();

            //QAT
            if (Properties.Settings.Default.QATLoaction == 0)
            {
                kryptonRibbon1.QATLocation = QATLocation.Above;
            }
            else if (Properties.Settings.Default.QATLoaction == 1)
            {
                kryptonRibbon1.QATLocation = QATLocation.Below;
            }


            //EdgeWebViewの初期化
            LoadEBDataFolder();
            


        }

        async Task LoadEBDataFolder()
        {
            try
            {
                var env = await CoreWebView2Environment.CreateAsync(null,Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"RTEditor", "WebView2"));

                await webView21.EnsureCoreWebView2Async(env);
            }
            catch
            {
                MessageBox.Show("このソフトウェアでは Edge WebView2 Runtime を使用しますがインストールが確認できないため起動できませんでした。Edge WebView2 Runtime のダウンロードサイトに移動します。", "RTEditor", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://go.microsoft.com/fwlink/p/?LinkId=2124703",
                    UseShellExecute = true
                });

                System.Windows.Forms.Application.Exit();
            }

        }



        //クリップボードの内容を監視して貼り付け方法の種類を変更する処理
        private void ClipBoradStatus_Tick(object sender, EventArgs e)
        {
            if (Clipboard.ContainsText())
            {
                kryptonContextMenuImageSelect2.ImageIndexStart = 0;
                kryptonContextMenuImageSelect2.ImageIndexEnd = 1;
            }
            else
            {
                kryptonContextMenuImageSelect2.ImageIndexStart = 1;
                kryptonContextMenuImageSelect2.ImageIndexEnd = 1;
            }
        }


        //保存ボタンをクリックしたときの処理
        //非同期処理で上書き保存を実行する
        private async void kryptonContextMenuItem4_Click(object sender, EventArgs e)
        {
            IsFormClosing = false;
            activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            fname = activeform.Text;

            activeform.Activate();
            string str = activeform.Text.Replace("*", "");
            // 既存のファイルパスが存在する場合はそのまま上書き保存
            if (File.Exists(str))
            {
                if (str.Contains(".rtf") | str.Contains(".txt.rtf"))
                {
                    richTextBox?.SaveFile(str, RichTextBoxStreamType.RichText);
                }
                else if (str.Contains(".txt") | str.Contains(".rtf.txt"))
                {
                    richTextBox?.SaveFile(str, RichTextBoxStreamType.PlainText);
                }

                // 保存が完了したらタイトルから*を外す
                activeform.Text = str;
            }
            // 既存のファイルパスが存在しない場合は名前を付けて保存ダイアログを表示して保存
            else
            {
                // 保存処理
                using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Title = "保存する場所を選択...", Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt", FilterIndex = Properties.Settings.Default.DefaultSaveMode + 1 })
                {
                    string str1 = activeform.Text.Replace("*", "");

                    if (File.Exists(str1) == true)
                    {
                        saveFileDialog.FileName = str1;
                    }
                    else
                    {
                        if (richTextBox?.Text.Length > 0)
                        {
                            saveFileDialog.FileName = richTextBox?.Lines[0];
                        }
                        else
                        {
                            saveFileDialog.FileName = "無題";
                        }
                    }

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            if (saveFileDialog.FilterIndex == 1)
                            {
                                richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.RichText);
                            }
                            else if (saveFileDialog.FilterIndex == 2)
                            {
                                richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
                            }
                            // 保存が完了したらタイトルから*を外す
                            activeform.Text = saveFileDialog.FileName;

                            //最近使用したドキュメントのファイルパス保存・表示処理
                            if (Properties.Settings.Default.RecentDocsPath.Contains(saveFileDialog.FileName) == false)
                            {
                                AddRecentDocs(saveFileDialog.FileName);
                                SaveRecentDocs(saveFileDialog.FileName);
                            }

                            //保存変更確認基準の更新
                            filepath = saveFileDialog.FileName;
                            ReadOnlyRTB.LoadFile(saveFileDialog.FileName);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("ファイルの保存中にエラーが発生しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }

                }
            }
        }

        //名前を付けて保存ボタンをクリックしたときの処理
        private void kryptonContextMenuItem5_Click(object sender, EventArgs e)
        {
            IsFormClosing = false;
            activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            fname = activeform.Text;

            // 保存処理
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Title = "保存する場所を選択...", Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt", FilterIndex = Properties.Settings.Default.DefaultSaveMode + 1 })
            {
                string str1 = activeform.Text.Replace("*", "");

                if (File.Exists(str1) == true)
                {
                    saveFileDialog.FileName = str1;
                }
                else
                {
                    if (richTextBox?.Text.Length > 0)
                    {
                        saveFileDialog.FileName = richTextBox?.Lines[0];
                    }
                    else
                    {
                        saveFileDialog.FileName = "無題";
                    }
                }

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog.FilterIndex == 1)
                        {
                            richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.RichText);
                        }
                        else if (saveFileDialog.FilterIndex == 2)
                        {
                            richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
                        }
                        // 保存が完了したらタイトルから*を外す
                        activeform.Text = saveFileDialog.FileName;

                        //最近使用したドキュメントのファイルパス保存・表示処理
                        if (Properties.Settings.Default.RecentDocsPath.Contains(saveFileDialog.FileName) == false)
                        {
                            AddRecentDocs(saveFileDialog.FileName);
                            SaveRecentDocs(saveFileDialog.FileName);
                        }

                        //保存変更確認基準の更新
                        filepath = saveFileDialog.FileName;
                        ReadOnlyRTB.LoadFile(saveFileDialog.FileName);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ファイルの保存中にエラーが発生しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

            }
        }

        //印刷ボタンをクリックしたときの処理
        private void kryptonContextMenuItem6_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            PrintDocument pd = new PrintDocument();
            pd.BeginPrint += Pd_BeginPrint;
            pd.PrintPage += Pd_PrintPage;
            printText = richTextBox;

            using (PrintDialog printDialog = new PrintDialog() { Document = pd, UseEXDialog = true, AllowSomePages = true})
            {
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    pd.PrinterSettings.PrinterName = printDialog.PrinterSettings.PrinterName;
                    pd.PrinterSettings.PrintRange = printDialog.PrinterSettings.PrintRange;
                    pd.DefaultPageSettings = printDialog.PrinterSettings.DefaultPageSettings;
                    pd.PrinterSettings.ToPage = printDialog.PrinterSettings.ToPage;
                    pd.Print();
                }

            }
        }

        //印刷プレビューをクリックしたときの処理
        private void kryptonContextMenuItem7_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            PrintDocument pd = new PrintDocument();
            pd.BeginPrint += Pd_BeginPrint;
            pd.PrintPage += Pd_PrintPage;
            printText = richTextBox;

            PrintPreviewDialog dlg = new PrintPreviewDialog();
            dlg.UseAntiAlias = true;
            dlg.Document = pd;
            dlg.Show();
            dlg.Activate();

        }

        private void kryptonContextMenuItem8_Click(object sender, EventArgs e)
        {

        }

        //選択したMdiフォームを閉じる
        private void kryptonContextMenuItem11_Click(object sender, EventArgs e)
        {
            if(this.MdiChildren.Length > 0)
            {
                activeform = this.ActiveMdiChild;
                activeform.Close();
            }
        }

        //すべてのドキュメントを閉じる処理
        private void kryptonContextMenuItem12_Click(object sender, EventArgs e)
        {
            
            while (this.MdiChildren.Length > 0)
            {
                IsFormClosing = false;
                activeform = this.ActiveMdiChild;
                activeform.Close();
                //すべてのMdiウィンドウを閉じた後、またはキャンセルされた場合、カウンターをリセットする
                if (this.MdiChildren.Length == 0 | result == DialogResult.Cancel)
                {
                    count = 0;
                    //ループを抜ける
                    break;
                }
            }
        }

        //選択したドキュメントまたはファイルを完全に削除する処理
        private void kryptonContextMenuItem20_Click(object sender, EventArgs e)
        {
            if (this.MdiChildren.Length > 0)
            {
                activeform = this.ActiveMdiChild;
                string filename = activeform.Text.Replace("*", "");
                if (File.Exists(filename) == true)
                {
                    if (Properties.Settings.Default.IsShowDeleteWarningTaskDialog == true)
                    {
                        Microsoft.WindowsAPICodePack.Dialogs.TaskDialog td = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog();
                        td.Caption = "ファイル削除確認";
                        td.InstructionText = "選択したドキュメントまたはファイルを完全に削除しようとしています";
                        td.Text = filename + "を完全に削除しようとしています。この操作はシステムのごみ箱には移動されず完全に削除されます。よろしいですか？";
                        td.Icon = TaskDialogStandardIcon.Warning;
                        td.StandardButtons = TaskDialogStandardButtons.Yes | TaskDialogStandardButtons.No;
                        td.FooterCheckBoxText = "次回削除時にこのメッセージを表示しない";
                        td.OwnerWindowHandle = this.Handle;
                        if (td.Show() == TaskDialogResult.Yes)
                        {

                            try
                            {
                                activeform.Text = "無題";
                                File.Delete(filename);
                                activeform.Dispose();
                                //最近使用したドキュメントをチェックし存在しないファイルがあれば項目から削除
                                kryptonRibbon1.RibbonAppButton.AppButtonRecentDocs.Clear();
                                ScanRecentDocs();

                                if (td.FooterCheckBoxChecked == true)
                                {
                                    Properties.Settings.Default.IsShowDeleteWarningTaskDialog = false;
                                    Properties.Settings.Default.Save();
                                }
                                else if (td.FooterCheckBoxChecked == false)
                                {
                                    Properties.Settings.Default.IsShowDeleteWarningTaskDialog = true;
                                    Properties.Settings.Default.Save();
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ファイルの削除中にエラーが発生しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                    else if (Properties.Settings.Default.IsShowDeleteWarningTaskDialog == false)
                    {
                        try
                        {
                            activeform.Text = "無題";
                            File.Delete(filename);
                            activeform.Dispose();
                            //最近使用したドキュメントをチェックし存在しないファイルがあれば項目から削除
                            kryptonRibbon1.RibbonAppButton.AppButtonRecentDocs.Clear();
                            ScanRecentDocs();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("ファイルの削除中にエラーが発生しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }

            }


        }


        //ドラッグアンドドロップでファイルを開く処理
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            int fileCount = 0;
            while (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                MdiFormCreate(sender, e);
                Form activeform = this.ActiveMdiChild;
                RichTextBox richTextBox = (RichTextBox)activeform.Controls[0];
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                filepath = files[fileCount];
                //ファイルをリッチテキストボックスに表示


                if (files[fileCount].Contains(".rtf") | files[fileCount].Contains(".txt.rtf"))
                {
                    richTextBox.LoadFile(files[fileCount], RichTextBoxStreamType.RichText);
                    ReadOnlyRTB.LoadFile(files[fileCount], RichTextBoxStreamType.RichText);
                }
                else if (files[fileCount].Contains(".txt") | files[fileCount].Contains(".rtf.txt"))
                {
                    richTextBox.LoadFile(files[fileCount], RichTextBoxStreamType.PlainText);
                    ReadOnlyRTB.LoadFile(files[fileCount], RichTextBoxStreamType.PlainText);
                }

                //最近使用したドキュメントのファイルパス保存・表示処理
                if (Properties.Settings.Default.RecentDocsPath.Contains(files[fileCount]) == false)
                {
                    AddRecentDocs(files[fileCount]);
                    SaveRecentDocs(files[fileCount]);
                }

                //タスクバーのジャンプリストに追加
                AddRecentFileJumplist(files[fileCount], jumpList);


                activeform.Text = files[fileCount];
                filepath = files[fileCount];
                fileCount++;
                if (fileCount == files.Length)
                {
                    fileCount = 0;
                    break;
                }

            }

        }

        //ドラッグアンドドロップで親ウィンドウにドロップしたときの処理
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void Form1_DragLeave(object sender, EventArgs e)
        {
        }

        //親ウィンドウを完全に閉じた時にクリップボードの内容の自動取得を中止する
        //また不必要なウィンドウオブジェクトを破棄する
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if(kryptonRibbon1.QATLocation == QATLocation.Above)
            {
                Properties.Settings.Default.QATLoaction = 0;
            }
            else if(kryptonRibbon1.QATLocation == QATLocation.Below)
            {
                Properties.Settings.Default.QATLoaction = 1;
            }

            if (kryptonRibbonQATButton1.Visible == true)
            {
                Properties.Settings.Default.QAT1Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT1Visible = false;
            }

            if (kryptonRibbonQATButton2.Visible == true)
            {
                Properties.Settings.Default.QAT2Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT2Visible = false;
            }

            if (kryptonRibbonQATButton3.Visible == true)
            {
                Properties.Settings.Default.QAT3Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT3Visible = false;
            }

            if (kryptonRibbonQATButton4.Visible == true)
            {
                Properties.Settings.Default.QAT4Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT4Visible = false;
            }

            if (kryptonRibbonQATButton5.Visible == true)
            {
                Properties.Settings.Default.QAT5Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT5Visible = false;
            }

            if (kryptonRibbonQATButton6.Visible == true)
            {
                Properties.Settings.Default.QAT6Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT6Visible = false;
            }

            if (kryptonRibbonQATButton7.Visible == true)
            {
                Properties.Settings.Default.QAT7Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT7Visible = false;
            }

            if (kryptonRibbonQATButton8.Visible == true)
            {
                Properties.Settings.Default.QAT8Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT8Visible = false;
            }
            Properties.Settings.Default.Save();


            Properties.Settings.Default.Save();
               

           RemoveClipboardFormatListener(Handle);

            GC.Collect();

            this.Dispose();
        }



        #region ファイルの貼り付け機能
        Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog;
        string selectedFilePath;
        int fileCount_Files;
        OpenFileDialog openFileDialog_Files;
        //ファイルの貼り付け方法を選択するダイアログを表示してから貼り付けを実行する処理
        private void kryptonRibbonGroupButton12_Click(object sender, EventArgs e)
        {
            using (openFileDialog_Files = new OpenFileDialog() { Filter = "すべてのファイル(*.*)|*.*", Multiselect = true, Title = "挿入するファイルを選択..." })
            {
                if (openFileDialog_Files.ShowDialog() == DialogResult.OK)
                {
                    fileCount_Files = 0;
                    selectedFilePath = openFileDialog_Files.FileNames[fileCount_Files];
                    taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog();
                    taskDialog.Caption = "ファイルの貼り付け方法確認";
                    taskDialog.InstructionText = "ファイルの貼り付け方法を選択してください";
                    taskDialog.Text = "選択したファイルをリッチテキストボックスに貼り付ける方法を選択してください。";
                    TaskDialogCommandLink pasteAsOLEDoc = new TaskDialogCommandLink("pasteAsText", "OLEオブジェクトとして貼り付け", "選択したファイルをOLEオブジェクトとして貼り付けます");
                    pasteAsOLEDoc.Click += pasteAsOLEDoc_Click;
                    taskDialog.Controls.Add(pasteAsOLEDoc);
                    TaskDialogCommandLink pasteAsText = new TaskDialogCommandLink("pasteAsText", "テキストとして貼り付け", "選択したファイルをファイルパステキストとして貼り付けます");
                    taskDialog.Controls.Add(pasteAsText);
                    pasteAsText.Click += pasteAsText_Click;
                    TaskDialogCommandLink cancelButton = new TaskDialogCommandLink("cancel", "キャンセル", "貼り付けをキャンセルします");
                    taskDialog.Controls.Add(cancelButton);
                    cancelButton.Click += cancelButton_Click;

                    taskDialog.OwnerWindowHandle = this.Handle;
                    taskDialog.Show();

                }
            }
        }

        //ファイルをOLEオブジェクトとして貼り付ける処理
        public void pasteAsOLEDoc_Click(object sender, EventArgs e)
        {
            taskDialog.Close();
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            fileCount_Files = 0;
            while (fileCount_Files < openFileDialog_Files.FileNames.Length)
            {
                var data = new DataObject();
                data.SetData(DataFormats.FileDrop, new string[] { openFileDialog_Files.FileNames[fileCount_Files] });

                Clipboard.SetDataObject(data, true);

                richTextBox.Paste();
                fileCount_Files++;
                if (fileCount_Files == openFileDialog_Files.FileNames.Length)
                {
                    break;
                }
            }


        }

        //ファイルをファイルパステキストとして貼り付ける処理
        public void pasteAsText_Click(object sender, EventArgs e)
        {
            taskDialog.Close();

            while (fileCount_Files < openFileDialog_Files.FileNames.Length)
            {
                Form activeform = this.ActiveMdiChild;
                richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

                richTextBox.AppendText(openFileDialog_Files.FileNames[fileCount_Files] + "\n");
                fileCount_Files++;
                if (fileCount_Files == openFileDialog_Files.FileNames.Length)
                {
                    break;
                }
            }
        }


        //貼り付け処理をキャンセルする
        public void cancelButton_Click(object sender, EventArgs e)
        {
            taskDialog.Close();
        }
        #endregion



        //通常の貼り付けボタンをクリックしたときの処理
        private void kryptonRibbonGroupButton1_Click(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Paste();
        }

        //クリップボードの内容を貼り付ける方法を選択するコンテキストメニューの処理
        private void kryptonContextMenuImageSelect2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            //書式付きテキストで貼り付け
            if (kryptonContextMenuImageSelect2.SelectedIndex == 0)
            {
                //テキストのみ貼り付け
                if (Clipboard.ContainsText())
                {
                    //そのまま貼り付ける
                    string fontName = kryptonRibbonGroupComboBox1.Text;
                    kryptonRibbonGroupComboBox1.Text = string.Empty;
                    richTextBox.Paste();
                    kryptonRibbonGroupComboBox1.Text = fontName;
                }

            }
            //プレーンテキストで貼り付け
            else if (kryptonContextMenuImageSelect2.SelectedIndex == 1)
            {
                //フォントをリボンのコンボボックスから取得しテキストのみ貼り付け
                if (Clipboard.ContainsText())
                {
                    string plainText = Clipboard.GetText();
                    //フォントはkryptonRibbonGroupComboBox1.Textのものを使いつつフォントスタイルとサイズは現行のまま使う
                    richTextBox.SelectionFont = new Font(kryptonRibbonGroupComboBox1.Text, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style);
                    richTextBox.AppendText(plainText);
                }
                else
                {
                    richTextBox.Paste();
                }
            }
            //ボタンクリックのみで動作するようにSelectedIndexをリセット
            kryptonContextMenuImageSelect2.SelectedIndex = -1;
        }

        private void kryptonContextMenuImageSelect2_TrackingImage(object sender, ImageSelectEventArgs e)
        {
        }

        //アプリケーションメニューでドキュメントを新規作成する処理
        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
            MdiFormCreate(sender, e);
        }

        //コピー処理
        private void kryptonRibbonGroupButton2_Click_1(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Copy();
        }


        //各リッチテキストのフォントスタイル状態等を確認するメソッド
        public void richTextBox_ScanFontStyle()
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            try
            {
                //フォントスタイルの種類
                //太字　kryptonRibbonGroupClusterButton1
                //斜体　kryptonRibbonGroupClusterButton2
                //下線　kryptonContextMenuItem23
                //打ち消し線　kryptonContextMenuItem24
                //上付き文字　kryptonRibbonGroupClusterButton4
                //下付き文字　kryptonRibbonGroupClusterButton5
                //文字色※使わない　kryptonRibbonGroupClusterColorButton1
                //マーカー色※使わない　kryptonRibbonGroupClusterColorButton2
                //フォント名称　kryptonRibbonGroupComboBox1
                //フォントサイズ　kryptonRibbonGroupComboBox2
                //左寄せ　kryptonRibbonGroupClusterButton8
                //中央寄せ　kryptonRibbonGroupClusterButton9
                //右寄せ　kryptonRibbonGroupClusterButton10
                //箇条書き　kryptonRibbonGroupClusterButton13

                //太字
                if (richTextBox.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                    toolStripButton1.Checked = true;
                    radialMenuItem1.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton1.Checked = false;
                    toolStripButton1.Checked = false;
                    radialMenuItem1.Checked = false;
                }
                //斜体
                if (richTextBox.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                    toolStripButton2.Checked = true;
                    radialMenuItem2.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton2.Checked = false;
                    toolStripButton2.Checked = false;
                    radialMenuItem2.Checked = false;
                }
                //下線
                if (richTextBox.SelectionFont.Underline)
                {
                    kryptonContextMenuItem23.Checked = true;
                    toolStripMenuItem2.Checked = true;
                    radialMenuItem3.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem23.Checked = false;
                    toolStripMenuItem2.Checked = false;
                    radialMenuItem3.Checked = false;
                }
                //打ち消し線
                if (richTextBox.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem24.Checked = true;
                    toolStripMenuItem3.Checked = true;
                    radialMenuItem4.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem24.Checked = false;
                    toolStripMenuItem3.Checked = false;
                    radialMenuItem4.Checked = false;
                }
                //標準文字確認
                if (richTextBox.SelectionCharOffset != 0)
                {
                    //上付き文字
                    if (richTextBox.SelectionCharOffset == 5)
                    {
                        kryptonRibbonGroupClusterButton4.Checked = true;
                        kryptonRibbonGroupClusterButton5.Checked = false;
                    }
                    //下付き文字
                    else if (richTextBox.SelectionCharOffset == -5)
                    {
                        kryptonRibbonGroupClusterButton5.Checked = true;
                        kryptonRibbonGroupClusterButton4.Checked = false;
                    }
                }
                else
                {
                    kryptonRibbonGroupClusterButton5.Checked = false;
                    kryptonRibbonGroupClusterButton4.Checked = false;
                }


                //フォント名称
                kryptonRibbonGroupComboBox1.Text = richTextBox.SelectionFont.Name;
                toolStripComboBox2.Text = richTextBox.SelectionFont.Name;

                //フォントサイズ
                kryptonRibbonGroupComboBox2.Text = richTextBox.SelectionFont.Size.ToString();
                toolStripComboBox3.Text = richTextBox.SelectionFont.Size.ToString();
                //左寄せ
                if (richTextBox.SelectionAlignment == HorizontalAlignment.Left)
                {
                    kryptonRibbonGroupClusterButton8.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton8.Checked = false;
                }
                //中央寄せ
                if (richTextBox.SelectionAlignment == HorizontalAlignment.Center)
                {
                    kryptonRibbonGroupClusterButton9.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton9.Checked = false;
                }
                //右寄せ
                if (richTextBox.SelectionAlignment == HorizontalAlignment.Right)
                {
                    kryptonRibbonGroupClusterButton10.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton10.Checked = false;
                }
                //箇条書き
                if (richTextBox.SelectionBullet == true)
                {
                    kryptonRibbonGroupClusterButton13.Checked = true;
                    toolStripButton5.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton13.Checked = false;
                    toolStripButton5.Checked = false;
                }

                //Undo・Redoできるか確認
                if (richTextBox.CanUndo == true)
                {
                    kryptonRibbonQATButton7.Enabled = true;
                    kryptonRibbonQATButton7.ToolTipTitle = "元に戻す(Ctrl+Z)";
                }
                else
                {
                    kryptonRibbonQATButton7.Enabled = false;
                    kryptonRibbonQATButton7.ToolTipTitle = "元に戻せません(Ctrl+Z)";
                }
                if (richTextBox.CanRedo == true)
                {
                    kryptonRibbonQATButton8.Enabled = true;
                    kryptonRibbonQATButton8.ToolTipTitle = "やり直す(Ctrl+Y)";

                }
                else
                {
                    kryptonRibbonQATButton8.Enabled = false;
                    kryptonRibbonQATButton8.ToolTipTitle = "やり直せません(Ctrl+Y)";
                }

                //読み取り用か書き込み対応か
                if (richTextBox.ReadOnly == false)
                {
                    kryptonRibbonGroupButton9.Checked = true;
                    kryptonRibbonGroupButton10.Checked = false;

                    閲覧ToolStripMenuItem.Checked = false;
                    編集ToolStripMenuItem.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupButton9.Checked = false;
                    kryptonRibbonGroupButton10.Checked = true;

                    閲覧ToolStripMenuItem.Checked = true;
                    編集ToolStripMenuItem.Checked = false;
                }

                //段落のインデントと間隔
                kryptonRibbonGroupNumericUpDown1.Value = richTextBox.SelectionIndent;
                kryptonRibbonGroupNumericUpDown2.Value = richTextBox.SelectionRightIndent;
                kryptonRibbonGroupNumericUpDown3.Value = richTextBox.SelectionCharOffset;

                //選択箇所の文字列の保護状態
                if (richTextBox.SelectionProtected == true)
                {
                    kryptonRibbonGroupButton38.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupButton38.Checked = false;
                }

                //リッチテキストボックスに左端の空白
                if (richTextBox.ShowSelectionMargin == true)
                {
                    kryptonRibbonGroupCheckBox1.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupCheckBox1.Checked = false;
                }

            }
            catch { }
        }

        //書式コピー用の各プロパティ
        public string fontName { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikeout { get; set; }
        public bool IsSuperscript { get; set; }
        public bool IsSubscript { get; set; }
        public bool IsNoSuperscriptAndSubscript { get; set; }
        public Color textColor { get; set; }
        public Color textHighlighter { get; set; }
        public float FontSize { get; set; }

        //書式コピー
        private void kryptonContextMenuItem28_Click(object sender, EventArgs e)
        {

            //書式情報を初期化
            fontName = string.Empty;
            IsBold = false;
            IsItalic = false;
            IsUnderline = false;
            IsStrikeout = false;
            IsSuperscript = false;
            IsSubscript = false;
            IsNoSuperscriptAndSubscript = false;
            textColor = Color.Empty;
            textHighlighter = Color.Empty;
            FontSize = 0;

            //選択中のリッチテキストボックスのテキストから書式情報を取得
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);

            fontName = richTextBox.SelectionFont.FontFamily.Name;
            //太字
            if (richTextBox.SelectionFont.Bold)
            {
                IsBold = true;
            }
            else
            {
                IsBold = false;
            }

            //斜体
            if (richTextBox.SelectionFont.Italic)
            {
                IsItalic = true;
            }
            else
            {
                IsItalic = false;
            }

            //下線
            if (richTextBox.SelectionFont.Underline)
            {
                IsUnderline = true;
            }
            else
            {
                IsUnderline = false;
            }

            //取り消し線
            if (richTextBox.SelectionFont.Strikeout)
            {
                IsStrikeout = true;
            }
            else
            {
                IsStrikeout = false;
            }

            //上付き文字
            if (richTextBox.SelectionCharOffset == 5)
            {
                IsSuperscript = true;
            }
            else
            {
                IsSuperscript = false;
            }

            //下付き文字
            if (richTextBox.SelectionCharOffset == -5)
            {
                IsSubscript = true;
            }
            else
            {
                IsSubscript = false;
            }

            //標準文字
            if (richTextBox.SelectionCharOffset == 0)
            {
                IsNoSuperscriptAndSubscript = true;
            }
            else
            {
                IsNoSuperscriptAndSubscript = false;
            }

            //文字色
            textColor = richTextBox.SelectionColor;
            //マーカー色
            textHighlighter = richTextBox.SelectionBackColor;
            //フォントサイズ
            FontSize = richTextBox.SelectionFont.Size;

            //初めて書式コピーを行う際、kryptonContextMenuItem30をクリックできるようにする
            kryptonContextMenuItem30.Enabled = true;
            書式貼り付けToolStripMenuItem.Enabled = true;
        }

        //書式貼り付け
        private void kryptonContextMenuItem30_Click(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);

            richTextBox.SelectionFont = new Font(fontName, FontSize, Font.Style);
            //太字
            if (IsBold == true)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont, richTextBox.SelectionFont.Style | FontStyle.Bold);
            }

            //斜体
            if (IsItalic == true)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont, richTextBox.SelectionFont.Style | FontStyle.Italic);
            }

            //下線
            if (IsUnderline == true)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont, richTextBox.SelectionFont.Style | FontStyle.Underline);
            }

            //取り消し線
            if (IsStrikeout == true)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont, richTextBox.SelectionFont.Style | FontStyle.Strikeout);
            }

            //標準文字確認
            if (IsNoSuperscriptAndSubscript != true)
            {
                //上付き文字
                if (IsSuperscript == true)
                {
                    richTextBox.SelectionCharOffset = 5;
                }
                //下付き文字
                else if (IsSubscript == true)
                {
                    richTextBox.SelectionCharOffset = -5;
                }
            }
            else
            {
                richTextBox.SelectionCharOffset = 0;
            }

            richTextBox.SelectionColor = textColor;
            richTextBox.SelectionBackColor = textHighlighter;

        }

        //切り取り処理
        private void kryptonRibbonGroupButton3_Click(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Cut();
        }

        //文字削除処理
        private void kryptonRibbonGroupButton4_Click(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectedText = string.Empty;
        }


        private void kryptonDockingManager1_FloatingWindowRemoved(object sender, FloatingWindowEventArgs e)
        {

        }

        private void kryptonDockingManager1_PageFloatingRequest(object sender, CancelUniqueNameEventArgs e)
        {



        }


        private void kryptonDockingManager1_PageCloseRequest(object sender, CloseRequestEventArgs e)
        {

        }


        private void kryptonHeaderGroup1_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void kryptonHeaderGroup1_DragLeave(object sender, EventArgs e)
        {

        }





        private void kryptonRibbon1_SelectedContextChanged(object sender, EventArgs e)
        {

            if (kryptonRibbon1.SelectedContext != "")
            {
                this.StateCommon.Header.Content.ShortText.TextH = PaletteRelativeAlign.Inherit;

            }
            else
            {
                this.StateCommon.Header.Content.ShortText.TextH = PaletteRelativeAlign.Center;
            }
        }

        //アプリケーションメニュー内のキーボードショートカットを使用されたときの処理
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            IsUsingMiniToolBar = false;
            miniToolBar1.Hide();

            //ドキュメントを閉じるホットキー
            if (e.Control && e.KeyCode == Keys.W)
            {
                kryptonContextMenuItem11_Click(sender, e);
            }

            //ドキュメントをすべて閉じるホットキー
            if (e.Control && e.Shift && e.KeyCode == Keys.W)
            {
                kryptonContextMenuItem12_Click(sender, e);
            }

            //ファイルを削除するホットキー
            if (e.Control && e.Shift && e.KeyCode == Keys.D)
            {
                kryptonContextMenuItem20_Click(sender, e);
            }

            //印刷するホットキー
            if (e.Control && e.KeyCode == Keys.P)
            {
                if (e.Control != e.Shift && e.KeyCode == Keys.P)
                {
                    kryptonContextMenuItem6_Click(sender, e);
                }
            }

            //印刷プレビューを表示するホットキー
            if (e.Control && e.Shift && e.KeyCode == Keys.P)
            {
                kryptonContextMenuItem7_Click(sender, e);

                Form activeForm = this.ActiveMdiChild;
                richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                string rtfTable =
                @"{\rtf1\ansi
                \trowd\cellx2000\cellx4000
                \intbl A\cell
                \row
                \trowd\cellx2000\cellx4000
                \intbl A\cell
                \row
                }";

                richTextBox.SelectedRtf = rtfTable;

            }

            //書式コピーと書式貼り付けのホットキー
            if (e.Control && e.Alt && e.KeyCode == Keys.C)
            {
                kryptonContextMenuItem28_Click(sender, e);
            }

            if (e.Control && e.Alt && e.KeyCode == Keys.V)
            {
                kryptonContextMenuItem30_Click(sender, e);
            }

        }

        #region リッチテキストボックスのUndo・Redo処理
        private void kryptonRibbonQATButton7_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (richTextBox.CanUndo == true)
            {
                richTextBox.Undo();
                kryptonRibbonQATButton7.Enabled = true;
                kryptonRibbonQATButton7.ToolTipTitle = "元に戻す(Ctrl+Z)";

            }
            else
            {
                kryptonRibbonQATButton7.Enabled = false;
                kryptonRibbonQATButton7.ToolTipTitle = "元に戻せません(Ctrl+Z)";
            }

        }

        private void kryptonRibbonQATButton8_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (richTextBox.CanRedo == true)
            {
                richTextBox.Redo();
                kryptonRibbonQATButton8.Enabled = true;
                kryptonRibbonQATButton8.ToolTipTitle = "やり直す(Ctrl+Y)";
            }
            else
            {
                kryptonRibbonQATButton8.Enabled = false;
                kryptonRibbonQATButton8.ToolTipTitle = "やり直せません(Ctrl+Y)";
            }
        }
        #endregion

        private void すべて選択ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectAll();
        }


        //リボンホームタブ

        //フローティングウィンドウ表示されたクリップボードパネルを閉じ、再度パネルとして表示する処理
        public void fm2_Closing(object sneder, FormClosingEventArgs e)
        {
            kryptonPanel11.Show();
            kryptonPanel11.Controls.Add(kryptonHeaderGroup1);
            buttonSpecHeaderGroup1.Visible = true;
            buttonSpecHeaderGroup2.Visible = true;
            splitter1.Visible = true;
        }

        //クリップボードパネルをフローティングウィンドウ表示するための処理
        KryptonForm fm2;
        private void buttonSpecHeaderGroup1_Click(object sender, EventArgs e)
        {
            kryptonPanel11.Hide();
            fm2 = new KryptonForm();
            fm2.Palette = kryptonPalette1;
            fm2.TopMost = true;
            fm2.ShowIcon = false;
            fm2.MinimizeBox = false;
            fm2.FormClosing += fm2_Closing;
            fm2.Controls.Add(kryptonHeaderGroup1);
            fm2.Show();
            buttonSpecHeaderGroup1.Visible = false;
            buttonSpecHeaderGroup2.Visible = false;
            splitter1.Visible = false;

        }

        //クリップボードパネルを隠す処理
        private void buttonSpecHeaderGroup2_Click(object sender, EventArgs e)
        {
            kryptonPanel11.Hide();
            splitter1.Visible = false;
            kryptonRibbonGroup1.DialogBoxLauncher = true;
        }


        //クリップボードパネルを開く処理
        private void kryptonRibbonGroup1_DialogBoxLauncherClick(object sender, EventArgs e)
        {
            if(this.MdiChildren.Length != 0)
            {
                splitter1.Visible = true;
                kryptonPanel11.Show();
                kryptonPanel11.Show();
                kryptonRibbonGroup1.DialogBoxLauncher = false;
            }
           
        }

        //クリップボードの履歴を自動で取得するWindowsAPI関数
        [DllImport("user32.dll", SetLastError = true)]
        private extern static void AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private extern static void RemoveClipboardFormatListener(IntPtr hwnd);

        private const int WM_CLIPBOARDUPDATE = 0x31D;

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
            {
                OnClipboardUpdate();
                m.Result = IntPtr.Zero;
            }
            else
                base.WndProc(ref m);
        }

        //クリップボードの内容を取得したテキストと画像を表示するオブジェクト
        private KryptonButton[] clipBoradPasteButton;
        private PictureBox[] pictureBox1;

        //クリップボードの内容を自動取得するかのbool型のフラグ
        public bool IsScanClipBorad { get; set; } = true;

        //SnippingToolオーバーレイが表示中か確かめるフラグ
        public bool IsShowSnippingToolOverlay { get; set; } = false;

        //クリップボードの内容が変更されたときの処理
        void OnClipboardUpdate()
        {
            if (IsScanClipBorad == true)
            {
                if (Clipboard.ContainsText())
                {
                    this.clipBoradPasteButton = new KryptonButton[1];
                    for (int i = 0; i < this.clipBoradPasteButton.Length; i++)
                    {
                        this.clipBoradPasteButton[i] = new KryptonButton();
                        this.clipBoradPasteButton[i].Dock = DockStyle.Top;
                        this.clipBoradPasteButton[i].Height = 100;
                        this.clipBoradPasteButton[i].Text = Clipboard.GetText();
                        this.clipBoradPasteButton[i].StateCommon.Content.ShortText.TextH = PaletteRelativeAlign.Near;
                        this.clipBoradPasteButton[i].StateCommon.Content.ShortText.TextV = PaletteRelativeAlign.Near;
                        this.clipBoradPasteButton[i].ButtonStyle = ButtonStyle.ButtonSpec;
                        this.clipBoradPasteButton[i].AutoSize = true;
                        this.clipBoradPasteButton[i].AutoSizeMode = AutoSizeMode.GrowAndShrink;
                        this.clipBoradPasteButton[i].Palette = kryptonPalette1;

                        this.clipBoradPasteButton[i].Click += new EventHandler(clipBoradPasteButton_Click);

                    }
                    kryptonPanel12.Controls.AddRange(clipBoradPasteButton);

                    IsShowSnippingToolOverlay = false;
                }
                else if (Clipboard.ContainsImage())
                {
                    this.pictureBox1 = new PictureBox[1];
                    for (int i = 0; i < this.pictureBox1.Length; i++)
                    {
                        this.pictureBox1[i] = new PictureBox();
                        this.pictureBox1[i].Image = Clipboard.GetImage();
                        this.pictureBox1[i].Dock = DockStyle.Top;
                        this.pictureBox1[i].Height = 100;

                        this.pictureBox1[i].MouseEnter += new EventHandler(pic1_MouseEnter);
                        this.pictureBox1[i].MouseLeave += new EventHandler(pic1_MouseLeave);
                        this.pictureBox1[i].Click += new EventHandler(pic1_Click);

                        this.pictureBox1[i].SizeMode = PictureBoxSizeMode.Zoom;
                        this.pictureBox1[i].BackColor = Color.Transparent;
                    }
                    kryptonPanel12.Controls.AddRange(pictureBox1);

                    if (IsShowSnippingToolOverlay == true)
                    {
                        Form activeform = this.ActiveMdiChild;
                        richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                        Thread.Sleep(400); //スニッピングツールの処理が完了するまで少し待機
                        richTextBox.Paste();
                        IsShowSnippingToolOverlay = false;
                        this.WindowState = FormWindowState.Normal;
                    }

                }
            }
        }


        #region 拡大画像表示用のツールチップを表示するための処理
        ClipBoradPictureToolTip clipBoradPictureToolTip = new ClipBoradPictureToolTip();

        System.Windows.Forms.Timer timer3 = new System.Windows.Forms.Timer();
        public void timer3_Tick(object ssender, EventArgs e)
        {
            clipBoradPictureToolTip.Location = Cursor.Position + new Size(10, 10);
        }

        public void pic1_MouseEnter(object sender, EventArgs e)
        {
            timer3.Interval = 10;
            timer3.Tick += timer3_Tick;
            timer3.Start();
            clipBoradPictureToolTip.image = ((PictureBox)sender).Image;
            clipBoradPictureToolTip.Location = Cursor.Position + new Size(10, 10);
            clipBoradPictureToolTip.Show();
        }

        public void pic1_MouseLeave(object sender, EventArgs e)
        {
            clipBoradPictureToolTip.Hide();
            timer3.Stop();
        }
        #endregion

        //クリップボードから取得されたテキストをリッチテキストボックスに貼り付ける処理
        public void clipBoradPasteButton_Click(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText(((KryptonButton)sender).Text);
        }

        //クリップボードから取得された画像をリッチテキストボックスに貼り付ける処理
        public void pic1_Click(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            Clipboard.SetImage(((PictureBox)sender).Image);
            richTextBox.Paste();
        }

        //自動取得したクリップボードの内容すべてクリアする処理
        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            kryptonPanel12.Controls.Clear();
        }


        //クリップボードの内容自動取得機能の有効・無効を切り替える処理
        private void kryptonCheckButton3_Click(object sender, EventArgs e)
        {
            if (kryptonCheckButton3.Checked == true)
            {
                IsScanClipBorad = true;
            }
            else
            {
                IsScanClipBorad = false;
            }

        }

        //異なるフォントの文字を選択しフォントスタイルを変更しようとすると例外が発生するのでtry-catchで異なるフォントの文字をスタイル変更を防ぐ
        private void richTextBox_SelectionChanged(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;

            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();


            richTextBox_ScanFontStyle();
            try
            {
                kryptonRibbonGroupComboBox1.Enabled = true;
                kryptonRibbonGroupComboBox2.Enabled = true;

                toolStripComboBox2.Enabled = true;
                toolStripComboBox3.Enabled = true;


                //フォント名称
                kryptonRibbonGroupComboBox1.Text = richTextBox.SelectionFont.Name;
                toolStripComboBox2.Text = richTextBox.SelectionFont.Name;
                //フォントサイズ
                kryptonRibbonGroupComboBox2.Text = richTextBox.SelectionFont.Size.ToString();
                toolStripComboBox3.Text = richTextBox.SelectionFont.Size.ToString();

            }
            catch
            {
                kryptonRibbonGroupComboBox1.Enabled = false;
                kryptonRibbonGroupComboBox2.Enabled = false;

                toolStripComboBox2.Enabled = false;
                toolStripComboBox3.Enabled = false;

                kryptonRibbonGroupComboBox1.Text = "(混在)";
                kryptonRibbonGroupComboBox2.Text = "(混在)";

                toolStripComboBox2.Text = "(混在)";
                toolStripComboBox3.Text = "(混在)";
            }

        }

        //すべてのリボンの各コントロールの有効状態の切り替えメソッド
        public void EnabledRibbonAllGroup(object sender, EventArgs e)
        {
            kryptonRibbonGroupLabel3.Enabled = true;
            kryptonRibbonGroupLabel2.Enabled = true;
            kryptonRibbonGroupLabel5.Enabled = true;

            閲覧ToolStripMenuItem.Enabled = true;
            編集ToolStripMenuItem.Enabled = true;

            kryptonPage1.Visible = true;
            kryptonPage5.Visible = true;
            kryptonRibbonGroup1.DialogBoxLauncher = true;

            //クイックアクセスツールバー(QAT1～QAT2までは無視)
            kryptonRibbonQATButton3.Enabled = true;
            kryptonRibbonQATButton4.Enabled = true;
            kryptonRibbonQATButton5.Enabled = true;
            kryptonRibbonQATButton6.Enabled = true;
            //Undo・Redoは無視
            //kryptonRibbonQATButton7.Enabled = true;
            //kryptonRibbonQATButton8.Enabled = true;

            //アプリケーションメニュー(新しいドキュメント、開く、復号化、設定、ヘルプ関連の機能は無視)
            //保存
            kryptonContextMenuItem4.Enabled = true;
            //名前を付けて保存
            kryptonContextMenuItem5.Enabled = true;
            //印刷
            kryptonContextMenuItem3.Enabled = true;
            //ドキュメントを閉じる
            kryptonContextMenuItem9.Enabled = true;
            //エクスプローラーで開く
            kryptonContextMenuItem21.Enabled = true;

            //リボンの各ボタン
            //ホームタブ
            //貼り付け
            kryptonRibbonGroupButton1.Enabled = true;
            //切り取り
            kryptonRibbonGroupButton3.Enabled = true;
            //コピー
            kryptonRibbonGroupButton2.Enabled = true;
            //削除
            kryptonRibbonGroupButton4.Enabled = true;
            //フォント選択コンボボックス
            kryptonRibbonGroupComboBox1.Enabled = true;
            //フォントサイズ変更コンボボックス
            kryptonRibbonGroupComboBox2.Enabled = true;
            //太字ボタン
            kryptonRibbonGroupClusterButton1.Enabled = true;
            //斜体ボタン
            kryptonRibbonGroupClusterButton2.Enabled = true;
            //下線・打ち消し線ボタン
            kryptonRibbonGroupClusterButton3.Enabled = true;
            //上付き文字ボタン
            kryptonRibbonGroupClusterButton4.Enabled = true;
            //下付き文字ボタン
            kryptonRibbonGroupClusterButton5.Enabled = true;
            //文字色ボタン
            kryptonRibbonGroupClusterColorButton1.Enabled = true;
            //蛍光ペンボタン
            kryptonRibbonGroupClusterColorButton2.Enabled = true;
            //数学記号ボタン
            kryptonRibbonGroupClusterButton15.Enabled = true;
            //文字種変換ボタン
            kryptonRibbonGroupClusterButton16.Enabled = true;
            //フォントサイズ拡大ボタン
            kryptonRibbonGroupClusterButton6.Enabled = true;
            //フォントサイズ縮小ボタン
            kryptonRibbonGroupClusterButton7.Enabled = true;
            //書式クリアボタン
            kryptonRibbonGroupClusterButton14.Enabled = true;
            //左寄せボタン
            kryptonRibbonGroupClusterButton8.Enabled = true;
            //中央寄せボタン
            kryptonRibbonGroupClusterButton9.Enabled = true;
            //右寄せボタン
            kryptonRibbonGroupClusterButton10.Enabled = true;
            //インデント減ボタン
            kryptonRibbonGroupClusterButton11.Enabled = true;
            //インデント増ボタン
            kryptonRibbonGroupClusterButton12.Enabled = true;
            //箇条書きボタン
            kryptonRibbonGroupClusterButton13.Enabled = true;
            //罫線ボタン
            kryptonRibbonGroupClusterButton17.Enabled = true;
            //ノートシールボタン
            kryptonRibbonGroupClusterButton18.Enabled = true;
            //スタイル選択ギャラリー
            kryptonRibbonGroupCustomControl3.Enabled = true;
            kryptonRibbonGroupGallery1.Enabled = true;
            //項目1ボタン
            kryptonRibbonGroupButton25.Enabled = true;
            //項目2ボタン
            kryptonRibbonGroupButton26.Enabled = true;
            //検索ボタン
            kryptonRibbonGroupButton5.Enabled = true;
            //置換ボタン
            kryptonRibbonGroupButton6.Enabled = true;
            //全選択ボタン
            kryptonRibbonGroupButton7.Enabled = true;
            //音声入力ボタン
            kryptonRibbonGroupButton8.Enabled = true;
            //編集モードボタン
            kryptonRibbonGroupButton9.Enabled = true;
            //閲覧モードボタン
            kryptonRibbonGroupButton10.Enabled = true;
            //Wikipediaサイドペインボタン(無視)
            //kryptonRibbonGroupButton16

            //挿入
            //新規ドキュメントボタン(無視)
            //kryptonRibbonGroupButton22
            //ドキュメント内容連結挿入ボタン
            kryptonRibbonGroupButton23.Enabled = true;
            //画像挿入ボタン
            kryptonRibbonGroupButton11.Enabled = true;
            //スクショボタン
            kryptonRibbonGroupButton27.Enabled = true;
            //OLEオブジェクト挿入ボタン
            kryptonRibbonGroupButton12.Enabled = true;
            //Webリンク挿入ボタン
            kryptonRibbonGroupButton13.Enabled = true;
            //テーブル挿入ボタン
            kryptonRibbonGroupButton21.Enabled = true;
            //グラフ挿入ボタン
            kryptonRibbonGroupButton24.Enabled = true;
            //OCRボタン
            kryptonRibbonGroupButton14.Enabled = true;
            //QRコード・バーコード挿入ボタン
            kryptonRibbonGroupButton15.Enabled = true;
            //現在日付挿入ボタン
            kryptonRibbonGroupButton17.Enabled = true;
            //現在時刻挿入ボタン
            kryptonRibbonGroupButton18.Enabled = true;
            //現在日付・時刻挿入ボタン
            kryptonRibbonGroupButton19.Enabled = true;
            //カスタムの日付・時刻挿入ボタン
            kryptonRibbonGroupButton20.Enabled = true;

            //レイアウト・校閲・リーダー
            //インデント左
            kryptonRibbonGroupNumericUpDown1.Enabled = true;
            kryptonRibbonGroupNumericUpDown1.ReadOnly = false;
            //インデント右
            kryptonRibbonGroupNumericUpDown2.Enabled = true;
            kryptonRibbonGroupNumericUpDown2.ReadOnly = false;
            //前後間隔
            kryptonRibbonGroupNumericUpDown3.Enabled = true;
            kryptonRibbonGroupNumericUpDown3.ReadOnly = false;
            //段落リセットボタン
            kryptonRibbonGroupButton37.Enabled = true;
            //スぺチェックボタン
            kryptonRibbonGroupButton29.Enabled = true;
            //音声読み上げボタン
            kryptonRibbonGroupButton30.Enabled = true;
            kryptonRibbonGroupButton30.Checked = false;
            //文字列の保護ボタン
            kryptonRibbonGroupButton38.Enabled = true;

            //AI機能
            //AIチャットパネルボタン(無視)
            //kryptonRibbonGroupButton50
            //校正ボタン
            kryptonRibbonGroupButton51.Enabled = true;
            //要約ボタン
            kryptonRibbonGroupButton56.Enabled = true;
            //丁寧語変換ボタン
            kryptonRibbonGroupButton57.Enabled = true;
            //調査ボタン
            kryptonRibbonGroupButton59.Enabled = true;
            //翻訳ボタン
            kryptonRibbonGroupButton60.Enabled = true;

            //表示
            //ズーム
            kryptonRibbonGroupButton39.Enabled = true;
            //ズームレベルリセットボタン
            kryptonRibbonGroupButton40.Enabled = true;
            //左端の段落選択用チェックボックス
            kryptonRibbonGroupCheckBox1.Enabled = true;
            //その他項目は無視する

            //開発者向け機能ボタン
            kryptonRibbonGroupButton52.Enabled = true;
            kryptonRibbonGroupButton52.Checked = false;
        }

        public void DisabledRibbonAllGroup(object sender, EventArgs e)
        {
            kryptonRibbonGroupLabel3.Enabled = false;
            kryptonRibbonGroupLabel2.Enabled = false;
            kryptonRibbonGroupLabel5.Enabled = false;

            閲覧ToolStripMenuItem.Enabled = false;
            編集ToolStripMenuItem.Enabled = false;

            kryptonPage1.Visible = false;
            kryptonPage5.Visible = false;
            kryptonRibbonGroup1.DialogBoxLauncher = false;

            //各ウィンドウとサイドバーを閉じる
            //リッチテキストエディタに関わるダイアログをすべて閉じる
            if (ocrDialog != null)
            {
                ocrDialog.Close();
            }
            if (barcodeAndQRCodeScanerDialog != null)
            {
                barcodeAndQRCodeScanerDialog.Close();
            }
            if (insertMathematicalSymbolsDialog != null)
            {
                insertMathematicalSymbolsDialog.Close();
            }
            if (AIPromptAnswarDialog.Instance != null)
            {
                AIPromptAnswarDialog.Instance.Close();
            }
            if (shortcutKeyDialog != null)
            {
                shortcutKeyDialog.Close();
            }
            if (rtbZoonFactorDialog != null)
            {
                rtbZoonFactorDialog.Close();
            }

            if (fm2 != null)
            {
                fm2.Close();
            }
            buttonSpecHeaderGroup2_Click(sender, e);
            buttonSpecAny4_Click(sender, e);

            //クイックアクセスツールバー(QAT1～QAT2までは無視)
            kryptonRibbonQATButton3.Enabled = false;
            kryptonRibbonQATButton4.Enabled = false;
            kryptonRibbonQATButton5.Enabled = false;
            kryptonRibbonQATButton6.Enabled = false;
            kryptonRibbonQATButton7.Enabled = false;
            kryptonRibbonQATButton8.Enabled = false;

            //アプリケーションメニュー(新しいドキュメント、開く、復号化、設定、ヘルプ関連の機能は無視)
            //保存
            kryptonContextMenuItem4.Enabled = false;
            //名前を付けて保存
            kryptonContextMenuItem5.Enabled = false;
            //印刷
            kryptonContextMenuItem3.Enabled = false;
            //ドキュメントを閉じる
            kryptonContextMenuItem9.Enabled = false;
            //エクスプローラーで開く
            kryptonContextMenuItem21.Enabled = false;

            //リボンの各ボタン
            //ホームタブ
            //貼り付け
            kryptonRibbonGroupButton1.Enabled = false;
            //切り取り
            kryptonRibbonGroupButton3.Enabled = false;
            //コピー
            kryptonRibbonGroupButton2.Enabled = false;
            //削除
            kryptonRibbonGroupButton4.Enabled = false;
            //フォント選択コンボボックス
            kryptonRibbonGroupComboBox1.Enabled = false;
            //フォントサイズ変更コンボボックス
            kryptonRibbonGroupComboBox2.Enabled = false;
            //太字ボタン
            kryptonRibbonGroupClusterButton1.Enabled = false;
            //斜体ボタン
            kryptonRibbonGroupClusterButton2.Enabled = false;
            //下線・打ち消し線ボタン
            kryptonRibbonGroupClusterButton3.Enabled = false;
            //上付き文字ボタン
            kryptonRibbonGroupClusterButton4.Enabled = false;
            //下付き文字ボタン
            kryptonRibbonGroupClusterButton5.Enabled = false;
            //文字色ボタン
            kryptonRibbonGroupClusterColorButton1.Enabled = false;
            //蛍光ペンボタン
            kryptonRibbonGroupClusterColorButton2.Enabled = false;
            //数学記号ボタン
            kryptonRibbonGroupClusterButton15.Enabled = false;
            //文字種変換ボタン
            kryptonRibbonGroupClusterButton16.Enabled = false;
            //フォントサイズ拡大ボタン
            kryptonRibbonGroupClusterButton6.Enabled = false;
            //フォントサイズ縮小ボタン
            kryptonRibbonGroupClusterButton7.Enabled = false;
            //書式クリアボタン
            kryptonRibbonGroupClusterButton14.Enabled = false;
            //左寄せボタン
            kryptonRibbonGroupClusterButton8.Enabled = false;
            //中央寄せボタン
            kryptonRibbonGroupClusterButton9.Enabled = false;
            //右寄せボタン
            kryptonRibbonGroupClusterButton10.Enabled = false;
            //インデント減ボタン
            kryptonRibbonGroupClusterButton11.Enabled = false;
            //インデント増ボタン
            kryptonRibbonGroupClusterButton12.Enabled = false;
            //箇条書きボタン
            kryptonRibbonGroupClusterButton13.Enabled = false;
            //罫線ボタン
            kryptonRibbonGroupClusterButton17.Enabled = false;
            //ノートシールボタン
            kryptonRibbonGroupClusterButton18.Enabled = false;
            //スタイル選択ギャラリー
            kryptonRibbonGroupCustomControl3.Enabled = false;
            kryptonRibbonGroupGallery1.Enabled = false;
            //項目1ボタン
            kryptonRibbonGroupButton25.Enabled = false;
            //項目2ボタン
            kryptonRibbonGroupButton26.Enabled = false;
            //検索ボタン
            kryptonRibbonGroupButton5.Enabled = false;
            //置換ボタン
            kryptonRibbonGroupButton6.Enabled = false;
            //全選択ボタン
            kryptonRibbonGroupButton7.Enabled = false;
            //音声入力ボタン
            kryptonRibbonGroupButton8.Enabled = false;
            //編集モードボタン
            kryptonRibbonGroupButton9.Enabled = false;
            //閲覧モードボタン
            kryptonRibbonGroupButton10.Enabled = false;
            //Wikipediaサイドペインボタン(無視)
            //kryptonRibbonGroupButton16

            //挿入
            //新規ドキュメントボタン(無視)
            //kryptonRibbonGroupButton22
            //ドキュメント内容連結挿入ボタン
            kryptonRibbonGroupButton23.Enabled = false;
            //画像挿入ボタン
            kryptonRibbonGroupButton11.Enabled = false;
            //スクショボタン
            kryptonRibbonGroupButton27.Enabled = false;
            //OLEオブジェクト挿入ボタン
            kryptonRibbonGroupButton12.Enabled = false;
            //Webリンク挿入ボタン
            kryptonRibbonGroupButton13.Enabled = false;
            //テーブル挿入ボタン
            kryptonRibbonGroupButton21.Enabled = false;
            //グラフ挿入ボタン
            kryptonRibbonGroupButton24.Enabled = false;
            //OCRボタン
            kryptonRibbonGroupButton14.Enabled = false;
            //QRコード・バーコード挿入ボタン
            kryptonRibbonGroupButton15.Enabled = false;
            //現在日付挿入ボタン
            kryptonRibbonGroupButton17.Enabled = false;
            //現在時刻挿入ボタン
            kryptonRibbonGroupButton18.Enabled = false;
            //現在日付・時刻挿入ボタン
            kryptonRibbonGroupButton19.Enabled = false;
            //カスタムの日付・時刻挿入ボタン
            kryptonRibbonGroupButton20.Enabled = false;

            //レイアウト・校閲・リーダー
            //インデント左
            kryptonRibbonGroupNumericUpDown1.Enabled = false;
            kryptonRibbonGroupNumericUpDown1.ReadOnly = true;
           //インデント右
           kryptonRibbonGroupNumericUpDown2.Enabled = false;
            kryptonRibbonGroupNumericUpDown2.ReadOnly = true;
            //前後間隔
            kryptonRibbonGroupNumericUpDown3.Enabled = false;
            kryptonRibbonGroupNumericUpDown3.ReadOnly = true;
            //段落リセットボタン
            kryptonRibbonGroupButton37.Enabled = false;
            //スぺチェックボタン
            kryptonRibbonGroupButton29.Enabled = false;
            //音声読み上げボタン
            kryptonRibbonGroupButton30.Enabled = false;
            kryptonRibbonGroupButton30.Checked = false;
            //文字列の保護ボタン
            kryptonRibbonGroupButton38.Enabled = false;

            //AI機能
            //AIチャットパネルボタン(無視)
            //kryptonRibbonGroupButton50
            //校正ボタン
            kryptonRibbonGroupButton51.Enabled = false;
            //要約ボタン
            kryptonRibbonGroupButton56.Enabled = false;
            //丁寧語変換ボタン
            kryptonRibbonGroupButton57.Enabled = false;
            //調査ボタン
            kryptonRibbonGroupButton59.Enabled = false;
            //翻訳ボタン
            kryptonRibbonGroupButton60.Enabled = false;

            //表示
            //ズーム
            kryptonRibbonGroupButton39.Enabled = false;
            //ズームレベルリセットボタン
            kryptonRibbonGroupButton40.Enabled = false;
            //左端の段落選択用チェックボックス
            kryptonRibbonGroupCheckBox1.Enabled = false;
            //その他項目は無視する

            //開発者向け機能ボタン
            kryptonRibbonGroupButton52.Enabled = false;
            kryptonRibbonGroupButton52.Checked = false;

            //リボンコンテキストの変更
            kryptonRibbon1.SelectedContext = null;
        }

        //リボンのフォント用のコンボボックスの値を変更したときの処理
        private void kryptonRibbonGroupComboBox1_TextUpdate(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            try
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new System.Drawing.Font(kryptonRibbonGroupComboBox1.Text, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style);
            }
            catch
            { }

            //ツールストリップのコンボボックスにも反映させる
            toolStripComboBox2.Text = kryptonRibbonGroupComboBox1.Text;
        }

        //リボンのフォントサイズ用のコンボボックスの値を変更したときの処理
        private void kryptonRibbonGroupComboBox2_TextUpdate(object sender, EventArgs e)
        {
            //String型からFloat型へTryParseし選択したフォントサイズのみ変更する
            if (float.TryParse(kryptonRibbonGroupComboBox2.Text, out float rtb_FontSize))
            {
                Form activeForm = this.ActiveMdiChild;
                richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new System.Drawing.Font(richTextBox.Font.Name, rtb_FontSize, richTextBox.SelectionFont.Style);
            }

            //ツールストリップのコンボボックスにも反映させる
            toolStripComboBox3.Text = kryptonRibbonGroupComboBox2.Text;
        }

        //フォントスタイルのリセットを行うためのメソッド
        public void FontStyleReset()
        {
            floatTryParse();
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            richTextBox.SelectionFont = new Font(kryptonRibbonGroupComboBox1.Text, f, FontStyle.Regular);
        }

        //フォントスタイルの種類
        //太字　kryptonRibbonGroupClusterButton1 toolStripButton1
        //斜体　kryptonRibbonGroupClusterButton2 toolStripButton2
        //下線　kryptonContextMenuItem23 toolStripMenuItem2
        //打ち消し線　kryptonContextMenuItem24 toolStripMenuItem2

        //太字ボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterButton1_Click(object sender, EventArgs e)
        {
            floatTryParse();
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            //太字が有効な場合
            if (kryptonRibbonGroupClusterButton1.Checked == true)
            {

                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Bold);

                //ボタンのCheckedをTrueにする
                kryptonRibbonGroupClusterButton1.Checked = true;
                toolStripButton1.Checked = true;
            }
            //太字が無効な場合
            else if (kryptonRibbonGroupClusterButton1.Checked == false)
            {
                //ボタンのCheckedをFalseにする
                kryptonRibbonGroupClusterButton1.Checked = false;
                toolStripButton1.Checked = false;

                //フォントスタイルをリセット
                FontStyleReset();

                //ほかのフォントスタイルが有効か確認する
                //斜体が有効な場合
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Italic);

                    //ボタンのCheckedをTrueにする
                    kryptonRibbonGroupClusterButton2.Checked = true;
                    toolStripButton2.Checked = true;
                }
                //下線が有効な場合
                if (kryptonContextMenuItem23.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Underline);

                    //コンテキストメニューアイテムのCheckedをTrueにする
                    kryptonContextMenuItem23.Checked = true;
                    toolStripMenuItem2.Checked = true;
                }
                //打ち消し線が有効な場合
                if (kryptonContextMenuItem24.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Strikeout);

                    //コンテキストメニューアイテムのCheckedをTrueにする
                    kryptonContextMenuItem24.Checked = true;
                    toolStripMenuItem3.Checked = true;
                }
            }
        }

        //斜体ボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterButton2_Click(object sender, EventArgs e)
        {
            floatTryParse();
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();


            //斜体が有効な場合
            if (kryptonRibbonGroupClusterButton2.Checked == true)
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Italic);

                //ボタンのCheckedをTrueにする
                kryptonRibbonGroupClusterButton2.Checked = true;
                toolStripButton2.Checked = true;
            }
            //斜体が無効な場合
            else if (kryptonRibbonGroupClusterButton2.Checked == false)
            {
                //ボタンのCheckedをFalseにする
                kryptonRibbonGroupClusterButton2.Checked = false;
                toolStripButton2.Checked = false;

                //フォントスタイルをリセット
                FontStyleReset();

                //ほかのフォントスタイルが有効か確認する
                //太字が有効な場合
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Bold);

                    //ボタンのCheckedをTrueにする
                    kryptonRibbonGroupClusterButton1.Checked = true;
                    toolStripButton1.Checked = true;
                }
                //下線が有効な場合
                if (kryptonContextMenuItem23.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Underline);

                    //コンテキストメニューアイテムのCheckedをTrueにする
                    kryptonContextMenuItem23.Checked = true;
                    toolStripMenuItem2.Checked = true;
                }
                //打ち消し線が有効な場合
                if (kryptonContextMenuItem24.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Strikeout);

                    //コンテキストメニューアイテムのCheckedをTrueにする
                    kryptonContextMenuItem24.Checked = true;
                    toolStripMenuItem3.Checked = true;
                }
            }
        }

        float f;
        public void floatTryParse()
        {
            if (float.TryParse(kryptonRibbonGroupComboBox2.Text, out float fontSize))
            {
                f = fontSize;
            }
        }

        //下線ボタンをクリックしたときの処理
        private void kryptonContextMenuItem23_Click(object sender, EventArgs e)
        {
            floatTryParse();
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            //下線が有効な場合
            if (kryptonContextMenuItem23.Checked == true)
            {

                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Underline);

                //コンテキストメニューアイテムのCheckedをTrueにする
                kryptonContextMenuItem23.Checked = true;
                toolStripMenuItem2.Checked = true;
            }
            //下線が無効な場合
            else if (kryptonContextMenuItem23.Checked == false)
            {
                //コンテキストメニューアイテムのCheckedをFalseにする
                kryptonContextMenuItem23.Checked = false;
                toolStripMenuItem2.Checked = false;

                //フォントスタイルをリセット
                FontStyleReset();

                //ほかのフォントスタイルが有効か確認する
                //太字が有効な場合
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Bold);

                    //ボタンのCheckedをTrueにする
                    kryptonRibbonGroupClusterButton1.Checked = true;
                    toolStripButton1.Checked = true;
                }
                //斜体が有効な場合
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Italic);

                    //ボタンのCheckedをTrueにする
                    kryptonRibbonGroupClusterButton2.Checked = true;
                    toolStripButton2.Checked = true;
                }
                //打ち消し線が有効な場合
                if (kryptonContextMenuItem24.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Strikeout);

                    //コンテキストメニューアイテムのCheckedをTrueにする
                    kryptonContextMenuItem24.Checked = true;
                    toolStripMenuItem3.Checked = true;
                }
            }
        }

        //打ち消し線ボタンをクリックしたときの処理
        private void kryptonContextMenuItem24_Click(object sender, EventArgs e)
        {
            floatTryParse();
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            //打ち消し線が有効な場合
            if (kryptonContextMenuItem24.Checked == true)
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Strikeout);

                //コンテキストメニューアイテムのCheckedをTrueにする
                kryptonContextMenuItem24.Checked = true;
                toolStripMenuItem3.Checked = true;
            }
            //打ち消し線が無効な場合
            else if (kryptonContextMenuItem24.Checked == false)
            {
                //コンテキストメニューアイテムのCheckedをFalseにする
                kryptonContextMenuItem24.Checked = false;
                toolStripMenuItem3.Checked = false;

                //フォントスタイルをリセット
                FontStyleReset();

                //ほかのフォントスタイルが有効か確認する
                //太字が有効な場合
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Bold);

                    //ボタンのCheckedをTrueにする
                    kryptonRibbonGroupClusterButton1.Checked = true;
                    toolStripButton1.Checked = true;
                }
                //斜体が有効な場合
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Italic);

                    //ボタンのCheckedをTrueにする
                    kryptonRibbonGroupClusterButton2.Checked = true;
                    toolStripButton2.Checked = true;
                }
                //下線が有効な場合
                if (kryptonContextMenuItem23.Checked == true)
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style | FontStyle.Underline);

                    //コンテキストメニューアイテムのCheckedをTrueにする
                    kryptonContextMenuItem23.Checked = true;
                    toolStripMenuItem2.Checked = true;
                }
            }
        }


        //上付き文字ボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterButton4_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            if (kryptonRibbonGroupClusterButton4.Checked == true)
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                kryptonRibbonGroupNumericUpDown3.Value = 5;
                richTextBox.SelectionCharOffset = 5;
                kryptonRibbonGroupClusterButton4.Checked = true;
                kryptonRibbonGroupClusterButton5.Checked = false;
            }
            else
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                kryptonRibbonGroupNumericUpDown3.Value = 0;
                richTextBox.SelectionCharOffset = 0;
                kryptonRibbonGroupClusterButton4.Checked = false;
                kryptonRibbonGroupClusterButton5.Checked = false;
            }
        }

        //下付き文字ボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterButton5_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            if (kryptonRibbonGroupClusterButton5.Checked == true)
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                kryptonRibbonGroupNumericUpDown3.Value = -5;
                richTextBox.SelectionCharOffset = -5;
                kryptonRibbonGroupClusterButton5.Checked = true;
                kryptonRibbonGroupClusterButton4.Checked = false;
            }
            else
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                kryptonRibbonGroupNumericUpDown3.Value = 0;
                richTextBox.SelectionCharOffset = 0;
                kryptonRibbonGroupClusterButton5.Checked = false;
                kryptonRibbonGroupClusterButton4.Checked = false;
            }
        }


        //文字色選択ボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterColorButton1_SelectedColorChanged(object sender, ColorEventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);

            richTextBox.SelectionColor = e.Color;

            kryptonRibbonGroupClusterColorButton1.SelectedColor = e.Color;

        }
        private void kryptonRibbonGroupClusterColorButton1_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);

            richTextBox.SelectionColor = kryptonRibbonGroupClusterColorButton1.SelectedColor;

            kryptonRibbonGroupClusterColorButton1.SelectedColor = kryptonRibbonGroupClusterColorButton1.SelectedColor;

        }

        //マーカー色選択ボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterColorButton2_SelectedColorChanged(object sender, ColorEventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);

            richTextBox.SelectionBackColor = e.Color;

            kryptonRibbonGroupClusterColorButton2.SelectedColor = e.Color;

        }

        private void kryptonRibbonGroupClusterColorButton2_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);

            richTextBox.SelectionBackColor = kryptonRibbonGroupClusterColorButton2.SelectedColor;

            kryptonRibbonGroupClusterColorButton2.SelectedColor = kryptonRibbonGroupClusterColorButton2.SelectedColor;

        }


        //フォントサイズを上げるボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterButton6_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            float f = float.Parse(kryptonRibbonGroupComboBox2.Text);
            if (f < 8 | f == 8)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 9, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "9";
            }
            else if (f < 9 | f == 9)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10";
            }
            else if (f < 10 | f == 10)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10.5f, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10.5";
            }
            if (f < 10.5f | f == 10.5f)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 11, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "11";
            }
            else if (f < 11 | f == 11)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "12";
            }
            else if (f < 12 | f == 12)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 14, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "14";
            }
            else if (f < 14 | f == 14)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 16, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "16";
            }
            else if (f < 16 | f == 16)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 18, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "18";
            }
            else if (f < 18 | f == 18)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 20, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "20";
            }
            else if (f < 20 | f == 20)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 22, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "22";
            }
            else if (f < 22 | f == 22)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 24, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "24";
            }
            else if (f < 24 | f == 24)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 26, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "26";
            }
            else if (f < 26 | f == 26)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 28, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "28";
            }
            else if (f < 28 | f == 28)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 36, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "36";
            }
            else if (f < 36 | f == 36)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 48, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "48";
            }
            else if (f < 48 | f == 48)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 72, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "72";
            }

            //ミニツールバーのフォントサイズコンボボックスにも反映させる
            toolStripComboBox3.Text = kryptonRibbonGroupComboBox2.Text;

        }

        //フォントサイズを下げるボタンをクリックしたときの処理
        private void kryptonRibbonGroupClusterButton7_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            float f = float.Parse(kryptonRibbonGroupComboBox2.Text);
            if (f > 72 | f == 72)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 48, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "48";
            }
            else if (f > 48 | f == 48)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 36, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "36";
            }
            else if (f > 36 | f == 36)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 28, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "28";
            }
            else if (f > 28 | f == 28)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 26, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "26";
            }
            else if (f > 26 | f == 26)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 24, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "24";
            }
            else if (f > 24 | f == 24)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 22, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "22";
            }
            else if (f > 22 | f == 22)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 20, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "20";
            }
            else if (f > 20 | f == 20)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 18, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "18";
            }
            else if (f > 18 | f == 18)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 16, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "16";
            }
            else if (f > 16 | f == 16)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 14, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "14";
            }
            else if (f > 14 | f == 14)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "12";
            }
            else if (f > 12 | f == 12)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 11, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "11";
            }
            else if (f > 11 | f == 11)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10.5f, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10.5";
            }
            else if (f > 10.5f | f == 10.5f)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10";
            }
            else if (f > 10 | f == 10)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 9, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "9";
            }
            else if (f > 9 | f == 9)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 8, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "8";
            }

            //ミニツールバーのフォントサイズコンボボックスにも反映させる
            toolStripComboBox3.Text = kryptonRibbonGroupComboBox2.Text;
        }

        //選択した文字のフォントスタイルを初期化したときの処理
        private void kryptonRibbonGroupClusterButton14_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionFont = new Font("Yu Gothic UI", 12, FontStyle.Regular);
            richTextBox.SelectionColor = Color.Black;
            richTextBox.SelectionBackColor = Color.Transparent;
        }

        //左寄せ
        private void kryptonRibbonGroupClusterButton8_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionAlignment = HorizontalAlignment.Left;

            kryptonRibbonGroupClusterButton8.Checked = true;
            kryptonRibbonGroupClusterButton9.Checked = false;
            kryptonRibbonGroupClusterButton10.Checked = false;
        }

        //中央寄せ
        private void kryptonRibbonGroupClusterButton9_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionAlignment = HorizontalAlignment.Center;

            kryptonRibbonGroupClusterButton8.Checked = false;
            kryptonRibbonGroupClusterButton9.Checked = true;
            kryptonRibbonGroupClusterButton10.Checked = false;
        }

        //右寄せ
        private void kryptonRibbonGroupClusterButton10_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionAlignment = HorizontalAlignment.Right;

            kryptonRibbonGroupClusterButton8.Checked = false;
            kryptonRibbonGroupClusterButton9.Checked = false;
            kryptonRibbonGroupClusterButton10.Checked = true;
        }

        //箇条書きを行ったときの処理
        private void kryptonRibbonGroupClusterButton13_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (kryptonRibbonGroupClusterButton13.Checked == true)
            {
                richTextBox.SelectionBullet = true;
            }
            else
            {
                richTextBox.SelectionBullet = false;
            }
        }


        //インデントを減らす
        private void kryptonRibbonGroupClusterButton11_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionIndent = richTextBox.SelectionIndent - 48;
        }


        //インデントを増やす
        private void kryptonRibbonGroupClusterButton12_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionIndent = richTextBox.SelectionIndent + 48;
        }

        //リボンのフォントダイアログ
        private void kryptonRibbonGroup2_DialogBoxLauncherClick(object sender, EventArgs e)
        {
            if(this.MdiChildren.Length != 0)
            {
                Form activeForm = this.ActiveMdiChild;
                richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

                int rtb_selectStart = richTextBox.SelectionStart;
                int rtb_selectLength = richTextBox.SelectionLength;

                using (FontDialog fontDialog = new FontDialog() { Font = richTextBox.SelectionFont, Color = richTextBox.SelectionColor, ShowColor = true, ShowEffects = true })
                {
                    if (fontDialog.ShowDialog() == DialogResult.OK)
                    {
                        richTextBox.Select(rtb_selectStart, rtb_selectLength);
                        kryptonRibbonGroupComboBox1.Text = fontDialog.Font.Name;
                        kryptonRibbonGroupComboBox2.Text = fontDialog.Font.Size.ToString();
                        richTextBox.SelectionFont = new Font(fontDialog.Font.Name, fontDialog.Font.Size, fontDialog.Font.Style);
                        if (fontDialog.Font.Bold)
                        {
                            kryptonRibbonGroupClusterButton1.Checked = true;
                        }
                        else
                        {
                            kryptonRibbonGroupClusterButton1.Checked = false;
                        }

                        if (fontDialog.Font.Italic)
                        {
                            kryptonRibbonGroupClusterButton2.Checked = true;
                        }
                        else
                        {
                            kryptonRibbonGroupClusterButton2.Checked = false;
                        }

                        if (fontDialog.Font.Underline)
                        {
                            kryptonContextMenuItem23.Checked = true;
                        }
                        else
                        {
                            kryptonContextMenuItem23.Checked = false;
                        }

                        if (fontDialog.Font.Strikeout)
                        {
                            kryptonContextMenuItem24.Checked = true;
                        }
                        else
                        {
                            kryptonContextMenuItem24.Checked = false;
                        }

                        kryptonRibbonGroupClusterColorButton1.SelectedColor = fontDialog.Color;
                    }
                }
            }
           
        }

        //フォントスタイルをすべて解除しデフォルトフォントに設定するメソッド
        public void ResetDefaultFont()
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, FontStyle.Regular);
            richTextBox.SelectionColor = Color.Black;
        }

        //文字のスタイルを変更したときの処理
        private void kryptonRibbonGroupGallery1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);

            ResetDefaultFont();
            //標準
            if (kryptonRibbonGroupGallery1.SelectedIndex == 0)
            {
                ResetDefaultFont();
            }
            //見出し
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 1)
            {
                richTextBox.SelectionFont = new Font("Yu Gothic UI", 14, FontStyle.Regular);
                richTextBox.SelectionColor = Color.FromArgb(206, 221, 224);
            }
            //モダン
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 2)
            {
                richTextBox.SelectionFont = new Font("Segoe UI", 12, FontStyle.Regular);
            }
            //強調
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 3)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, FontStyle.Bold);
            }
            //ノート
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 4)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, FontStyle.Italic);
            }
            //取り消し
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 5)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, FontStyle.Strikeout);
            }

            richTextBox_ScanFontStyle();
        }



        #region 音声認識機能
        //音声認識の開始・停止ボタン
        //
        SpeechRecognitionEngine sr;
        private void kryptonRibbonGroupButton8_Click(object sender, EventArgs e)
        {
            if (this.MdiChildren.Length > 0)
            {
                if (kryptonRibbonGroupButton8.Checked == true)
                {
                    sr = new SpeechRecognitionEngine();
                    richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    sr.LoadGrammar(new DictationGrammar());
                    sr.SpeechRecognized += sr_SpeechRecognized;
                    sr.SetInputToDefaultAudioDevice();
                    sr.RecognizeAsync(RecognizeMode.Multiple);
                    var player = new SoundPlayer(@"C:\Windows\Media\Speech On.wav");
                    player.Play();
                }
                else if (kryptonRibbonGroupButton8.Checked == false)
                {
                    kryptonRibbonGroupButton8.Checked = false;
                    sr.RecognizeAsyncStop();
                    sr.RecognizeAsyncCancel();
                    sr.Dispose();
                    var player = new SoundPlayer(@"C:\Windows\Media\Speech Off.wav");
                    player.Play();
                }
            }
            else
            {
                kryptonRibbonGroupButton8.Checked = false;
            }



        }

        public void sr_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (this.MdiChildren.Length > 0)
            {
                //認識結果をリッチテキストボックスに追加
                activeform = this.ActiveMdiChild;
                richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                richTextBox.Text += (e.Result.Text);
                //特定の音声コマンドに応じた処理
                if (e.Result.Text == "改行")
                {
                    richTextBox.Text = richTextBox.Text.Replace("改行", "");
                    richTextBox.Text += Environment.NewLine;
                }
                if (e.Result.Text == "スペース")
                {
                    richTextBox.Text = richTextBox.Text.Replace("スペース", "");
                    richTextBox.Text += " ";
                }
                else if (e.Result.Text == "終了")
                {
                    richTextBox.Text = richTextBox.Text.Replace("終了", "");
                    kryptonRibbonGroupButton8.Checked = false;
                    sr.RecognizeAsyncStop();
                    sr.RecognizeAsyncCancel();
                    sr.Dispose();
                    var player = new SoundPlayer(@"C:\Windows\Media\Speech Off.wav");
                    player.Play();

                }
            }
            else
            {
                kryptonRibbonGroupButton8.Checked = false;

            }
        }

        #endregion

        #region 文字検索と置換機能

        public string Keyword;

        int index;
        int _lastSearchIndex = 0;
        string s;

        //検索用
        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            Keyword = kryptonTextBox1.Text;

            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            string word = Keyword;
            s = Keyword;

            //ラジオボタンのチェックに応じて検索方法を変更する
            RichTextBoxFinds option = RichTextBoxFinds.None;
            if (kryptonRadioButton1.Checked == true)
            {
                option |= RichTextBoxFinds.None;
            }
            else if (kryptonRadioButton2.Checked == true)
            {
                option |= RichTextBoxFinds.MatchCase;
            }
            else if (kryptonRadioButton3.Checked == true)
            {
                option |= RichTextBoxFinds.WholeWord;
            }

            // カーソル位置から検索
            index = richTextBox.Find(word, _lastSearchIndex, option);

            if (index == -1)
            {
                // 先頭から探し直す（ラップ検索）
                index = richTextBox.Find(word, 0, option);

                if (index == -1)
                {
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialog td = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog();
                    td.InstructionText = "検索した文字列はこのドキュメントの項目には含まれていませんでした";
                    td.Text = "別の検索内容を入力するか別のドキュメントを検索してみてください。";
                    td.Caption = "検索結果";
                    td.DefaultButton = TaskDialogDefaultButton.Ok;
                    td.OwnerWindowHandle = this.Handle;
                    td.Icon = TaskDialogStandardIcon.None;
                    td.Show();
                    return;

                }
            }

            // 見つかった場所を選択
            richTextBox.Select(index, word.Length);
            richTextBox.Focus();
            richTextBox.ScrollToCaret();

            string result = "";
            //検索内容を履歴として保存
            foreach (string s in kryptonListBox1.Items)
            {
                result += s + Environment.NewLine;

            }

            if (result.Contains(kryptonTextBox1.Text) == false)
            {
                MessageBox.Show(result);
                kryptonListBox1.Items.Add(kryptonTextBox1.Text);
                kryptonListBox1.SelectedIndex = kryptonListBox1.TopIndex;
            }

            // 次回検索の開始位置
            _lastSearchIndex = index + word.Length;
        }

        private void buttonSpecAny4_Click(object sender, EventArgs e)
        {
            splitter2.Visible = false;
            kryptonNavigator1.Visible = false;


            kryptonRibbonGroupButton5.Checked = false;
            kryptonRibbonGroupButton6.Checked = false;
            kryptonRibbonGroupButton16.Checked = false;
            kryptonRibbonGroupButton50.Checked = false;
        }

        private void kryptonNavigator1_SelectedPageChanged(object sender, EventArgs e)
        {
            if (kryptonNavigator1.SelectedPage == kryptonPage1)
            {
                kryptonRibbonGroupButton5.Checked = true;
                kryptonRibbonGroupButton6.Checked = false;
                kryptonRibbonGroupButton16.Checked = false;
                kryptonRibbonGroupButton50.Checked = false;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage5)
            {
                kryptonRibbonGroupButton5.Checked = false;
                kryptonRibbonGroupButton6.Checked = true;
                kryptonRibbonGroupButton16.Checked = false;
                kryptonRibbonGroupButton50.Checked = false;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage6)
            {
                kryptonRibbonGroupButton5.Checked = false;
                kryptonRibbonGroupButton6.Checked = false;
                kryptonRibbonGroupButton16.Checked = true;
                kryptonRibbonGroupButton50.Checked = false;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage8)
            {
                kryptonRibbonGroupButton5.Checked = false;
                kryptonRibbonGroupButton6.Checked = false;
                kryptonRibbonGroupButton16.Checked = false;
                kryptonRibbonGroupButton50.Checked = true;
            }
            //検索インデックスをリセット
            index = 0;
            _lastSearchIndex = 0;
        }
        private void kryptonRibbonGroupButton5_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton5.Checked == true)
            {
                splitter2.Visible = true;
                kryptonNavigator1.Visible = true;


                kryptonRibbonGroupButton5.Checked = true;
                kryptonRibbonGroupButton6.Checked = false;
                kryptonRibbonGroupButton16.Checked = false;
                kryptonRibbonGroupButton50.Checked = false;

                kryptonNavigator1.SelectedPage = kryptonPage1;
            }
            else
            {
                buttonSpecAny4_Click(sender, e);
            }

        }

        private void kryptonRibbonGroupButton6_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton6.Checked == true)
            {
                splitter2.Visible = true;
                kryptonNavigator1.Visible = true;


                kryptonRibbonGroupButton5.Checked = false;
                kryptonRibbonGroupButton6.Checked = true;
                kryptonRibbonGroupButton16.Checked = false;
                kryptonRibbonGroupButton50.Checked = false;

                kryptonNavigator1.SelectedPage = kryptonPage5;
            }
            else
            {
                buttonSpecAny4_Click(sender, e);
            }
        }

        private void buttonSpecAny3_Click(object sender, EventArgs e)
        {
            kryptonTextBox1.ResetText();
        }

        private void kryptonTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                kryptonButton5_Click(sender, e); return;
            }
        }


        private void kryptonButton4_Click_1(object sender, EventArgs e)
        {
            kryptonListBox1.Items.Clear();
        }

        private void kryptonListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            kryptonTextBox1.ResetText();
            kryptonTextBox1.Text = kryptonListBox1.GetItemText(kryptonListBox1.SelectedItem);
        }

        //置換用の検索ボタン
        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            Keyword = kryptonTextBox2.Text;

            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            string word = Keyword;

            //ラジオボタンのチェックに応じて検索方法を変更する
            RichTextBoxFinds option = RichTextBoxFinds.None;
            if (kryptonRadioButton4.Checked == true)
            {
                option |= RichTextBoxFinds.None;
            }
            else if (kryptonRadioButton5.Checked == true)
            {
                option |= RichTextBoxFinds.MatchCase;
            }
            else if (kryptonRadioButton6.Checked == true)
            {
                option |= RichTextBoxFinds.WholeWord;
            }

            // カーソル位置から検索
            index = richTextBox.Find(word, _lastSearchIndex, option);

            if (index == -1)
            {
                // 先頭から探し直す（ラップ検索）
                index = richTextBox.Find(word, 0, option);

                if (index == -1)
                {
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialog td = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog();
                    td.InstructionText = "検索した文字列はこのドキュメントの項目には含まれていませんでした";
                    td.Text = "別の検索内容を入力するか別のドキュメントを検索してみてください。";
                    td.Caption = "検索結果";
                    td.DefaultButton = TaskDialogDefaultButton.Ok;
                    td.OwnerWindowHandle = this.Handle;
                    td.Icon = TaskDialogStandardIcon.None;
                    td.Show();
                    return;

                }
            }

            //検索内容を履歴として保存
            kryptonListBox1.Items.Add(kryptonTextBox1.Text);
            kryptonListBox1.SelectedIndex = kryptonListBox1.TopIndex;

            // 次回検索の開始位置
            _lastSearchIndex = index + word.Length;
        }

        int _searchStart = 0;

        public string FindText => kryptonTextBox2.Text;
        public string ReplaceText => kryptonTextBox3.Text;

        private int FindNext(string word, RichTextBoxFinds options)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            int index = richTextBox.Find(word, _searchStart, options);

            if (index == -1)
            {
                // 先頭に戻って再検索（折り返し）
                index = richTextBox.Find(word, 0, options);
                if (index == -1)
                    return -1;
            }

            richTextBox.Select(index, word.Length);
            richTextBox.ScrollToCaret();

            _searchStart = index + word.Length;
            return index;
        }


        private void ReplaceOne(string find, string replace, RichTextBoxFinds opt)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            // まず見つける
            int index = FindNext(find, opt);
            if (index == -1)
            {
                Microsoft.WindowsAPICodePack.Dialogs.TaskDialog td = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog();
                td.InstructionText = "検索した文字列はこのドキュメントの項目には含まれていませんでした";
                td.Text = "別の検索内容を入力するか別のドキュメントを検索してみてください。";
                td.Caption = "検索結果";
                td.DefaultButton = TaskDialogDefaultButton.Ok;
                td.OwnerWindowHandle = this.Handle;
                td.Icon = TaskDialogStandardIcon.None;
                td.Show();
                return;
            }

            // 選択されている部分を置換
            richTextBox.SelectedText = replace;

            // 置換後の位置から検索を続ける
            _searchStart = index + replace.Length;
        }

        private void ReplaceAll(string find, string replace, RichTextBoxFinds opt)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            int count = 0;
            _searchStart = 0;

            using (Ookii.Dialogs.WinForms.ProgressDialog progressDialog = new ProgressDialog())
            {
                progressDialog.WindowTitle = "処理中...";
                progressDialog.Text = "ドキュメント全文の置換を実行中です";
                progressDialog.Description = "この処理には時間がかかります。しばらくお待ちください...";
                progressDialog.DoWork += ProgressDialog_DoWork;
                progressDialog.ShowCancelButton = false;
                //ダイアログを表示
                progressDialog.Show();
            }

            while (true)
            {
                int index = richTextBox.Find(find, _searchStart, opt);
                if (index == -1)
                    break;

                richTextBox.Select(index, find.Length);
                richTextBox.SelectedText = replace;

                _searchStart = index + replace.Length;
                count++;
            }

            Microsoft.WindowsAPICodePack.Dialogs.TaskDialog td = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog();
            td.InstructionText = $"{count} 件の項目をすべて置換しました";
            td.Text = $"置換前の文字:{find}\n置換後の文字:{replace}\n置換した項目数:{count}";
            td.Caption = "置換完了";
            td.DefaultButton = TaskDialogDefaultButton.Ok;
            td.OwnerWindowHandle = this.Handle;
            td.Icon = TaskDialogStandardIcon.None;
            td.Show();
        }


        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            //ラジオボタンのチェックに応じて検索方法を変更する
            RichTextBoxFinds option = RichTextBoxFinds.None;
            if (kryptonRadioButton4.Checked == true)
            {
                option |= RichTextBoxFinds.None;
            }
            else if (kryptonRadioButton5.Checked == true)
            {
                option |= RichTextBoxFinds.MatchCase;
            }
            else if (kryptonRadioButton6.Checked == true)
            {
                option |= RichTextBoxFinds.WholeWord;
            }

            ReplaceOne(kryptonTextBox2.Text, kryptonTextBox3.Text, option);
        }

        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            //ラジオボタンのチェックに応じて検索方法を変更する
            RichTextBoxFinds option = RichTextBoxFinds.None;
            if (kryptonRadioButton4.Checked == true)
            {
                option |= RichTextBoxFinds.None;
            }
            else if (kryptonRadioButton5.Checked == true)
            {
                option |= RichTextBoxFinds.MatchCase;
            }
            else if (kryptonRadioButton6.Checked == true)
            {
                option |= RichTextBoxFinds.WholeWord;
            }

            ReplaceAll(kryptonTextBox2.Text, kryptonTextBox3.Text, option);
        }
        #endregion

        //ドキュメント内のテキストをすべて選択したときの処理
        private void kryptonRibbonGroupButton7_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectAll();
        }

        //編集用
        bool IsReadOnly = false;
        private void kryptonRibbonGroupButton9_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            richTextBox.ReadOnly = false;

            kryptonRibbonGroupButton9.Checked = true;
            kryptonRibbonGroupButton10.Checked = false;

            閲覧ToolStripMenuItem.Checked = false;
            編集ToolStripMenuItem.Checked = true;

            IsReadOnly = false;

        }

        //閲覧用
        private void kryptonRibbonGroupButton10_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            richTextBox.ReadOnly = true;

            kryptonRibbonGroupButton9.Checked = false;
            kryptonRibbonGroupButton10.Checked = true;

            閲覧ToolStripMenuItem.Checked = true;
            編集ToolStripMenuItem.Checked = false;

            IsReadOnly = true;
        }

        //リボン挿入タブ

        //WebサイトのURLを貼り付ける処理
        private void kryptonRibbonGroupButton13_Click(object sender, EventArgs e)
        {
            using (AddWebLinkDialog addWebLinkDialog = new AddWebLinkDialog())
            {
                addWebLinkDialog.ShowDialog();
                if (addWebLinkDialog.DialogResult == DialogResult.OK)
                {
                    Form activeform = this.ActiveMdiChild;
                    richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    richTextBox.AppendText(addWebLinkDialog.WebLinkURL);

                }
            }

        }
        //画像の貼り付け機能
        public void kryptonRibbonGroupButton11_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "PNGファイル(*.png)|*.png|JPG・JPEGファイル(*.jpg;*.jpeg)|*.jpg;*.jpeg|BMPファイル(*.bmp)|*.bmp|GIFファイル(*.gif)|*.gif|すべての画像ファイル(*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif", Multiselect = true, Title = "使用する画像ファイルを選択..." })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    int fileCount_Image = 0;
                    while (openFileDialog.FileNames.Length > 0)
                    {
                        Form activeform = this.ActiveMdiChild;
                        richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                        Clipboard.SetImage(new Bitmap(openFileDialog.FileNames[fileCount_Image]));
                        richTextBox.Paste();
                        fileCount_Image++;
                        if (fileCount_Image == openFileDialog.FileNames.Length)
                        {
                            break;
                        }
                    }

                }
            }
        }


        //テーブルの追加
        private void kryptonRibbonGroupButton21_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string rtfTable = @"{\rtf1\ansi\deff0 {\fonttbl {\f0 Courier New;}}
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
\trowd\cellx2000\cellx4000\cellx6000
\intbl 1 \cell\intbl 2 \cell\intbl 3 \cell\row
}";
            richTextBox.Rtf = rtfTable;
        }

        //表示中のフォームのテーマ適用
        public void ApplyFormThemes(Form FormName, KryptonPalette KryptonPalleteName)
        {
            try
            {
                //表示中のフォームのテーマ適用するため
                //まずフォームとKryptonPalleteが存在するか確かめる
                if (FormName != null && FormName.Visible != false && KryptonPalleteName != null)
                {
                    //各kryptonPalletteのアクセスレベルをPublicにする必要がある
                    //フォームがあればフォームのkryptonPaletteの値を変更しテーマを適用する
                    //テーマ適用

                    //既定のテーマ
                    if (Properties.Settings.Default.Theme == "Global")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Global;
                    }
                    //Professional-Systemテーマ
                    else if (Properties.Settings.Default.Theme == "ProfessionalSystem")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.ProfessionalSystem;
                    }
                    //Professional-Office2003テーマ
                    else if (Properties.Settings.Default.Theme == "ProfessionalOffice2003")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.ProfessionalOffice2003;
                    }
                    //Office2007Blueテーマ
                    else if (Properties.Settings.Default.Theme == "Office2007Blue")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                    }
                    //Office2007Silverテーマ
                    else if (Properties.Settings.Default.Theme == "Office2007Silver")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                    }
                    //Office2007Blackテーマ
                    else if (Properties.Settings.Default.Theme == "Office2007Black")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                    }
                    //Office2010Blueテーマ
                    else if (Properties.Settings.Default.Theme == "Office2010Blue")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                    }
                    //Office2010Silverテーマ
                    else if (Properties.Settings.Default.Theme == "Office2010Silver")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                    }
                    //Office2010Blackテーマ
                    else if (Properties.Settings.Default.Theme == "Office2010Black")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                    }
                    //SparkleBlueテーマ
                    else if (Properties.Settings.Default.Theme == "SparkleBlue")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.SparkleBlue;
                    }
                    //SparkleOrangeテーマ
                    else if (Properties.Settings.Default.Theme == "SparkleOrange")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.SparkleOrange;
                    }
                    //SparklePurpleテーマ
                    else if (Properties.Settings.Default.Theme == "SparklePurple")
                    {
                        KryptonPalleteName.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.SparklePurple;
                    }

                    if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == true)
                    {
                        KryptonPalleteName.AllowFormChrome = ComponentFactory.Krypton.Toolkit.InheritBool.False;
                    }
                    else if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == false)
                    {
                        KryptonPalleteName.AllowFormChrome = ComponentFactory.Krypton.Toolkit.InheritBool.True;
                    }

                }
            }
            catch
            { }

        }

        //設定ダイアログの表示と設定の保存
        private void kryptonContextMenuItem32_Click(object sender, EventArgs e)
        {
            using (SettingsDialog settingsDialog = new SettingsDialog())
            {
                //設定画面表示中はすべての各フォームを最小化する
                if (ocrDialog != null)
                {
                    ocrDialog.WindowState = FormWindowState.Minimized;
                }
                if (barcodeAndQRCodeScanerDialog != null)
                {
                    barcodeAndQRCodeScanerDialog.WindowState = FormWindowState.Minimized;
                }
                if (insertMathematicalSymbolsDialog != null)
                {
                    insertMathematicalSymbolsDialog.WindowState = FormWindowState.Minimized;
                }
                if(AIPromptAnswarDialog.Instance != null)
                {
                    AIPromptAnswarDialog.Instance.WindowState = FormWindowState.Minimized;
                }
                if(shortcutKeyDialog != null)
                {
                    shortcutKeyDialog.WindowState = FormWindowState.Minimized;
                }
                if (rtbZoonFactorDialog != null)
                {
                    rtbZoonFactorDialog.WindowState = FormWindowState.Minimized;
                }
                if (tutorialDialog != null)
                {
                    tutorialDialog.WindowState = FormWindowState.Minimized;
                }

                if (kryptonRibbonQATButton1.Visible == true)
                {
                    Properties.Settings.Default.QAT1Visible = true;
                }
                else 
                {
                    Properties.Settings.Default.QAT1Visible = false;
                }

                if (kryptonRibbonQATButton2.Visible == true)
                {
                    Properties.Settings.Default.QAT2Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT2Visible = false;
                }

                if (kryptonRibbonQATButton3.Visible == true)
                {
                    Properties.Settings.Default.QAT3Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT3Visible = false;
                }

                if (kryptonRibbonQATButton4.Visible == true)
                {
                    Properties.Settings.Default.QAT4Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT4Visible = false;
                }

                if (kryptonRibbonQATButton5.Visible == true)
                {
                    Properties.Settings.Default.QAT5Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT5Visible = false;
                }

                if (kryptonRibbonQATButton6.Visible == true)
                {
                    Properties.Settings.Default.QAT6Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT6Visible = false;
                }

                if (kryptonRibbonQATButton7.Visible == true)
                {
                    Properties.Settings.Default.QAT7Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT7Visible = false;
                }

                if (kryptonRibbonQATButton8.Visible == true)
                {
                    Properties.Settings.Default.QAT8Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT8Visible = false;
                }
                Properties.Settings.Default.Save();

                settingsDialog.ShowDialog();
                AppLoadSettings();
                //最近使用したドキュメントをチェックし存在しないファイルがあれば項目から削除
                kryptonRibbon1.RibbonAppButton.AppButtonRecentDocs.Clear();
                ScanRecentDocs();

                //各フォームのテーマ適用
                if (ocrDialog != null)
                {
                    ApplyFormThemes(ocrDialog, ocrDialog.kryptonPalette1);
                }
                if (barcodeAndQRCodeScanerDialog != null)
                {
                    ApplyFormThemes(barcodeAndQRCodeScanerDialog, barcodeAndQRCodeScanerDialog.kryptonPalette1);
                }
                if (insertMathematicalSymbolsDialog != null)
                {
                    ApplyFormThemes(insertMathematicalSymbolsDialog, insertMathematicalSymbolsDialog.kryptonPalette1);
                }
                if (AIPromptAnswarDialog.Instance != null)
                {
                    ApplyFormThemes(AIPromptAnswarDialog.Instance, AIPromptAnswarDialog.Instance.kryptonPalette1);
                }
                if (shortcutKeyDialog != null)
                {
                    ApplyFormThemes(shortcutKeyDialog, shortcutKeyDialog.kryptonPalette1);
                }
                if (rtbZoonFactorDialog != null)
                {
                    ApplyFormThemes(rtbZoonFactorDialog, rtbZoonFactorDialog.kryptonPalette1);
                }
                if (tutorialDialog != null)
                {
                    ApplyFormThemes(tutorialDialog, tutorialDialog.kryptonPalette1);
                }

            }

        }

        //文字をすべて大文字にする処理
        private void kryptonContextMenuItem36_Click(object sender, EventArgs e)
        {
            ToUpperAll();
        }

        //文字をすべて小文字にする処理
        private void kryptonContextMenuItem37_Click(object sender, EventArgs e)
        {
            ToLowerAll();
        }

        //文字をすべて大文字に変換するメソッド
        private void ToUpperAll()
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            var rtb = richTextBox;
            int selStart = rtb.SelectionStart;
            int selLength = rtb.SelectionLength;

            for (int i = 0; i < rtb.TextLength; i++)
            {
                rtb.Select(i, 1);
                string ch = rtb.SelectedText;
                if (ch.Length == 1 && char.IsLetter(ch[0]))
                    rtb.SelectedText = ch.ToUpper();
            }

            rtb.Select(selStart, selLength);
        }

        //文字をすべて小文字にする処理
        private void ToLowerAll()
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            var rtb = richTextBox;
            int selStart = rtb.SelectionStart;
            int selLength = rtb.SelectionLength;

            for (int i = 0; i < rtb.TextLength; i++)
            {
                rtb.Select(i, 1);
                string ch = rtb.SelectedText;
                if (ch.Length == 1 && char.IsLetter(ch[0]))
                    rtb.SelectedText = ch.ToLower();
            }

            rtb.Select(selStart, selLength);
        }

        //数式・記号挿入ダイアログの表示
        InsertMathematicalSymbolsDialog insertMathematicalSymbolsDialog;
        private void kryptonRibbonGroupClusterButton15_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupClusterButton15.Checked == true)
            {
                kryptonRibbonGroupClusterButton15.Checked = true;
                insertMathematicalSymbolsDialog = new InsertMathematicalSymbolsDialog();
                insertMathematicalSymbolsDialog.Show();
            }
            else
            {
                kryptonRibbonGroupClusterButton15.Checked = false;
                if (insertMathematicalSymbolsDialog.Visible == true)
                {
                    insertMathematicalSymbolsDialog.Close();
                }
            }
        }
        OCRDialog ocrDialog;
        private void kryptonRibbonGroupButton14_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton14.Checked == true)
            {
                ocrDialog = new OCRDialog();
                ocrDialog.Show();
            }
            else
            {
                ocrDialog.Close();
            }

        }

        //複数ファイルの内容を連結して貼り付ける機能
        private void kryptonRibbonGroupButton23_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Title = "開くファイルを選択...", Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt|リッチテキストファイルと書式なしテキストファイル(*.rtf;*.txt)|*.rtf;*.txt", Multiselect = true, ReadOnlyChecked = true })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Form activeform = this.ActiveMdiChild;
                    richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    RichTextBox copyOnlyRTB = new RichTextBox();

                    int i = 0;
                    while (i < openFileDialog.FileNames.Length)
                    {
                        copyOnlyRTB.ResetText();
                        copyOnlyRTB.LoadFile(openFileDialog.FileNames[i]);
                        copyOnlyRTB.SelectAll();
                        copyOnlyRTB.Copy();
                        copyOnlyRTB.DeselectAll();
                        richTextBox.Paste();
                        richTextBox.AppendText(Environment.NewLine);
                        i++;
                        if (i == openFileDialog.FileNames.Length)
                        {
                            copyOnlyRTB.Dispose();
                            break;
                        }
                    }

                }
            }
        }

        //バーコード・QRコードスキャナー機能
        BarcodeAndQRCodeScanerDialog barcodeAndQRCodeScanerDialog;
        private void kryptonRibbonGroupButton15_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton15.Checked == true)
            {
                barcodeAndQRCodeScanerDialog = new BarcodeAndQRCodeScanerDialog();
                barcodeAndQRCodeScanerDialog.Show();
            }
            else
            {
                barcodeAndQRCodeScanerDialog.Close();
            }

        }

        //日付と時刻挿入機能
        private void kryptonRibbonGroupButton17_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            DateTime dateTime = DateTime.Now;
            richTextBox.AppendText(dateTime.ToString("yyyy年M月d日"));
            if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday)
                richTextBox.AppendText("(日曜日)");
            else if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
                richTextBox.AppendText("(月曜日)");
            else if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday)
                richTextBox.AppendText("(火曜日)");
            else if (DateTime.Now.DayOfWeek == DayOfWeek.Wednesday)
                richTextBox.AppendText("(水曜日)");
            else if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday)
                richTextBox.AppendText("(木曜日)");
            else if (DateTime.Now.DayOfWeek == DayOfWeek.Friday)
                richTextBox.AppendText("(金曜日)");
            else if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                richTextBox.AppendText("(土曜日)");
        }

        //時刻挿入機能
        private void kryptonRibbonGroupButton18_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();

            DateTime dateTime = DateTime.Now;
            richTextBox.AppendText(dateTime.ToString("tt HH:mm"));
        }

        //日付と時刻を連続で挿入する機能
        private void kryptonRibbonGroupButton19_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton17_Click(sender, e);
            richTextBox.AppendText(" ");
            kryptonRibbonGroupButton18_Click(sender, e);
        }

        //日付と時刻挿入ダイアログ
        private void kryptonRibbonGroupButton20_Click(object sender, EventArgs e)
        {
            using (InsertCastumDateAndTimeDialog insertCastumDateAndTimeDialog = new InsertCastumDateAndTimeDialog())
            {
                insertCastumDateAndTimeDialog.ShowDialog();
                if (insertCastumDateAndTimeDialog.DialogResult == DialogResult.OK)
                {
                    Form activeForm = this.ActiveMdiChild;
                    richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    richTextBox.AppendText(insertCastumDateAndTimeDialog.Date);
                }
            }

        }

        //ズーム機能のトラックバー処理
        private void trackBarItem1_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                activeform = this.ActiveMdiChild;
                richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                richTextBox.ZoomFactor = kryptonTrackBar1.Value / 5F;
                int i = kryptonTrackBar1.Value * 5;
                kryptonLabel2.Text = i.ToString() + "%";
            }
            catch
            { }
        }

        private void kryptonTrackBar1_MouseEnter(object sender, EventArgs e)
        {
        }

        private void kryptonTrackBar1_MouseLeave(object sender, EventArgs e)
        {
        }

        //Webブラウザ表示機能のリボンボタン処理
        private void kryptonRibbonGroupButton16_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton16.Checked == true)
            {
                splitter2.Visible = true;
                kryptonNavigator1.Visible = true;


                kryptonRibbonGroupButton5.Checked = false;
                kryptonRibbonGroupButton6.Checked = false;
                kryptonRibbonGroupButton16.Checked = true;
                kryptonRibbonGroupButton50.Checked = false;

                kryptonNavigator1.SelectedPage = kryptonPage6;
            }
            else
            {
                buttonSpecAny4_Click(sender, e);
            }
        }

        //Webページの戻る処理
        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            if (webView21.CanGoBack)
            {
                webView21.GoBack();
            }

        }

        //Webページの進む処理
        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if (webView21.CanGoForward)
            {
                webView21.GoForward();
            }
        }

        //Webページの再読み込み処理
        private void buttonSpecAny9_Click(object sender, EventArgs e)
        {
            webView21.Refresh();
        }

        //既定のブラウザでWebページを開く処理
        private void buttonSpecAny10_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(webView21.Source.ToString());
        }

        //WebView2の読み込み完了時の処理
        private void webView21_ContentLoading(object sender, Microsoft.Web.WebView2.Core.CoreWebView2ContentLoadingEventArgs e)
        {
            kryptonTextBox4.Text = webView21.Source.ToString();
        }

        private void kryptonRibbonGroupButton24_Click(object sender, EventArgs e)
        {

        }

        //グラフ挿入機能
        private void kryptonRibbonGroupButton24_Click_1(object sender, EventArgs e)
        {
            using (InsertGraphDialog insertGraphDialog = new InsertGraphDialog())
            {
                insertGraphDialog.ShowDialog();
                if (insertGraphDialog.DialogResult == DialogResult.OK)
                {
                    Form activeForm = this.ActiveMdiChild;
                    richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
                    richTextBox.Paste();
                }
            }
        }

        //SnippingToolを使って画像をリッチテキストボックスに貼り付ける機能
        private void kryptonRibbonGroupButton27_Click(object sender, EventArgs e)
        {
            try
            {
                this.WindowState = FormWindowState.Minimized;
                Thread.Sleep(500); // ウィンドウが最小化されるのを待つ
                // Windows 11 Snipping Tool 起動
                Process.Start(new ProcessStartInfo("ms-screenclip:") { UseShellExecute = true });
                IsShowSnippingToolOverlay = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Snipping Tool の起動に失敗しました: " + ex.Message);
            }
        }


        //リボンTTSタブ

        //TTS読み上げ機能
        SpeechSynthesizer speechSynthesizer = new SpeechSynthesizer();

        //TTSリボンコンテキストの表示・非表示
        private void kryptonRibbonGroupButton30_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton52.Checked = false;
            if (kryptonRibbonGroupButton30.Checked == true)
            {
                speechSynthesizer.SpeakCompleted += speechSynthesizer_SpeakCompleted;

                kryptonRibbon1.SelectedContext = "TTSReaderContext";
                kryptonRibbon1.SelectedTab = kryptonRibbonTab7;

            }
            else
            {
                kryptonRibbon1.SelectedContext = string.Empty;
                //TTSの停止処理
                speechSynthesizer.Pause();
                speechSynthesizer.SpeakAsyncCancelAll();
            }
        }

        //TTS再生・一時停止ボタンの処理
        private void kryptonRibbonGroupButton32_Click(object sender, EventArgs e)
        {
            try
            {
                if (kryptonRibbonGroupButton32.TextLine1 == "停止中")
                {
                    statusStripLabel1.Text = "読み上げ中...";

                    kryptonRibbonGroupButton32.ImageLarge = Properties.Resources.Play;
                    kryptonRibbonGroupButton32.ImageSmall = Properties.Resources.Play;
                    kryptonRibbonGroupButton32.TextLine1 = "再生中";

                    //既定の音声
                    if (kryptonRibbonGroupComboBox3.SelectedIndex == 0)
                    {
                        speechSynthesizer.SelectVoiceByHints(VoiceGender.NotSet);
                    }
                    //女性の音声
                    else if (kryptonRibbonGroupComboBox3.SelectedIndex == 1)
                    {
                        speechSynthesizer.SelectVoiceByHints(VoiceGender.Female);
                    }
                    //男性の音声
                    else if (kryptonRibbonGroupComboBox3.SelectedIndex == 2)
                    {
                        speechSynthesizer.SelectVoiceByHints(VoiceGender.Male);
                    }
                    //中性的な音声
                    else if (kryptonRibbonGroupButton30.Checked == false)
                    {
                        speechSynthesizer.SelectVoiceByHints(VoiceGender.Neutral);
                    }

                    //TTSの開始処理
                    Form activeform = this.ActiveMdiChild;
                    richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    speechSynthesizer.Resume();
                    if (richTextBox.SelectionLength != 0)
                    {
                        string spText = richTextBox.SelectedText;
                        //改行コード削除してからTTSを開始する
                        string a = spText.Replace("\n", "");
                        speechSynthesizer.SpeakAsync(a);
                    }
                    else
                    {
                        string spText = richTextBox.Text;
                        //改行コード削除してからTTSを開始する
                        string a = spText.Replace("\n", "");
                        ; speechSynthesizer.SpeakAsync(a);
                    }


                }
                else
                {
                    statusStripLabel1.Text = "一時停止中";

                    kryptonRibbonGroupButton32.ImageLarge = Properties.Resources.Pause;
                    kryptonRibbonGroupButton32.ImageSmall = Properties.Resources.Pause;
                    kryptonRibbonGroupButton32.TextLine1 = "停止中";

                    //TTSの停止処理
                    speechSynthesizer.Pause();

                }
            }
            catch
            {
                kryptonRibbonGroupButton32.ImageLarge = Properties.Resources.Pause;
                kryptonRibbonGroupButton32.ImageSmall = Properties.Resources.Pause;
                kryptonRibbonGroupButton32.TextLine1 = "停止中";

                //TTSの停止処理
                speechSynthesizer.Pause();
            }

        }


        //TTS読み上げ完了時の処理
        private async void speechSynthesizer_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            kryptonRibbonGroupButton32.ImageLarge = Properties.Resources.Pause;
            kryptonRibbonGroupButton32.ImageSmall = Properties.Resources.Pause;
            kryptonRibbonGroupButton32.TextLine1 = "停止中";

            //TTSの停止処理
            speechSynthesizer.Pause();
            speechSynthesizer.SpeakAsyncCancelAll();

            statusStripLabel1.Text = "読み終わりました";
            await Task.Delay(3000);
            statusStripLabel1.Text = "準備完了";
        }

        //TTSミュートボタンの処理
        private void kryptonRibbonGroupButton35_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton35.Checked == true)
            {
                kryptonRibbonGroupTrackBar1.Enabled = false;
                speechSynthesizer.Volume = 0;
                kryptonRibbonGroupLabel9.TextLine1 = "0%(ミュート)";
            }
            else
            {
                kryptonRibbonGroupTrackBar1.Enabled = true;
                speechSynthesizer.Volume = kryptonRibbonGroupTrackBar1.Value;
                kryptonRibbonGroupLabel9.TextLine1 = kryptonRibbonGroupTrackBar1.Value.ToString() + "%";
            }

        }

        //TTS音量調整トラックバーの処理
        private void kryptonRibbonGroupTrackBar1_ValueChanged(object sender, EventArgs e)
        {
            speechSynthesizer.Volume = kryptonRibbonGroupTrackBar1.Value;
            kryptonRibbonGroupLabel9.TextLine1 = kryptonRibbonGroupTrackBar1.Value.ToString() + "%";
        }

        //TTSを一時停止しリセットする処理
        private void kryptonRibbonGroupButton31_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton32.ImageLarge = Properties.Resources.Pause;
            kryptonRibbonGroupButton32.ImageSmall = Properties.Resources.Pause;
            kryptonRibbonGroupButton32.TextLine1 = "停止中";

            //TTSの停止処理
            speechSynthesizer.Pause();
            speechSynthesizer.SpeakAsyncCancelAll();
        }

        private void kryptonRibbon1_SelectedTabChanged(object sender, EventArgs e)
        {
            statusStripLabel1.Text = "準備完了";
        }

        //TTS話速調整トラックバーの処理
        private void kryptonRibbonGroupTrackBar2_ValueChanged(object sender, EventArgs e)
        {
            speechSynthesizer.Rate = kryptonRibbonGroupTrackBar2.Value;
            kryptonRibbonGroupLabel10.TextLine1 = kryptonRibbonGroupTrackBar2.Value.ToString();

            if (kryptonRibbonGroupTrackBar2.Value == 0)
            {
                kryptonRibbonGroupLabel10.TextLine1 = "0(デフォルト)";
            }
        }

        //TTS話速リセットボタンの処理
        private void kryptonRibbonGroupButton36_Click(object sender, EventArgs e)
        {
            speechSynthesizer.Rate = 0;
            kryptonRibbonGroupTrackBar2.Value = 0;
            kryptonRibbonGroupLabel10.TextLine1 = "0(デフォルト)";
        }

        //WAVファイル出力ボタンの処理
        private void kryptonRibbonGroupButton33_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "WAVファイル (*.wav)|*.wav", Title = "WAVファイルを保存する場所を選択..." })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (SpeechSynthesizer saveOnlySpeechSynthesizer = new SpeechSynthesizer())
                        {

                            //既定の音声
                            if (kryptonRibbonGroupComboBox3.SelectedIndex == 0)
                            {
                                saveOnlySpeechSynthesizer.SelectVoiceByHints(VoiceGender.NotSet);
                            }
                            //女性の音声
                            else if (kryptonRibbonGroupComboBox3.SelectedIndex == 1)
                            {
                                saveOnlySpeechSynthesizer.SelectVoiceByHints(VoiceGender.Female);
                            }
                            //男性の音声
                            else if (kryptonRibbonGroupComboBox3.SelectedIndex == 2)
                            {
                                saveOnlySpeechSynthesizer.SelectVoiceByHints(VoiceGender.Male);
                            }
                            //中性的な音声
                            else if (kryptonRibbonGroupButton30.Checked == false)
                            {
                                saveOnlySpeechSynthesizer.SelectVoiceByHints(VoiceGender.Neutral);
                            }

                            saveOnlySpeechSynthesizer.Volume = speechSynthesizer.Volume;
                            saveOnlySpeechSynthesizer.Rate = speechSynthesizer.Rate;

                            saveOnlySpeechSynthesizer.SetOutputToDefaultAudioDevice();
                            saveOnlySpeechSynthesizer.SetOutputToWaveFile(saveFileDialog.FileName);
                            saveOnlySpeechSynthesizer.Speak(richTextBox.Text);
                            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "WAVファイル出力完了", InstructionText = "WAVファイルの出力が完了しました", Text = "ファイルの保存先:" + saveFileDialog.FileName, OwnerWindowHandle = this.Handle })
                            {
                                taskDialog.Show();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "WAVファイルの出力が正しく完了しませんでした", Text = "内容:" + ex.Message, OwnerWindowHandle = this.Handle, Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning })
                        {
                            taskDialog.Show();
                        }
                    }

                }
            }
        }

        //TTSリボンタブを閉じる処理
        private void kryptonRibbonGroupButton34_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton30.Checked = false;
            kryptonRibbon1.SelectedContext = string.Empty;
            //TTSの停止処理
            speechSynthesizer.Pause();
            speechSynthesizer.SpeakAsyncCancelAll();
        }

        //レイアウトタブ

        //左インデントの設定
        private void kryptonRibbonGroupNumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if(this.MdiChildren.Length > 0)
            {
                Form activeForm = this.ActiveMdiChild;
                richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

                richTextBox.SelectionIndent = (int)kryptonRibbonGroupNumericUpDown1.Value;
            }

        }

        //右インデントの設定
        private void kryptonRibbonGroupNumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if(this.MdiChildren.Length > 0)
            {
                Form activeForm = this.ActiveMdiChild;
                richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

                richTextBox.SelectionRightIndent = (int)kryptonRibbonGroupNumericUpDown2.Value;
            }

        }

        //上下間隔の設定
        private void kryptonRibbonGroupNumericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if(this.MdiChildren.Length > 0)
            {
                Form activeForm = this.ActiveMdiChild;
                richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
                if (richTextBox.SelectionLength == 0)
                {
                    richTextBox.SelectionCharOffset = (int)kryptonRibbonGroupNumericUpDown3.Value;
                }
                else
                {
                    richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                    richTextBox.SelectionCharOffset = (int)kryptonRibbonGroupNumericUpDown3.Value;
                }
            }

        }

        //選択部分の文字の保護設定
        private void kryptonRibbonGroupButton38_Click(object sender, EventArgs e)
        {
            richTextBox.ShowSelectionMargin = true;
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (kryptonRibbonGroupButton38.Checked == true)
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionProtected = true;
            }
            else
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionProtected = false;
            }

        }

        //インデントとをリセットするメソッド
        public void IndentReset()
        {
            kryptonRibbonGroupNumericUpDown1.Value = 0;
            kryptonRibbonGroupNumericUpDown2.Value = 0;
            kryptonRibbonGroupNumericUpDown3.Value = 0;
        }

        //文字の上下位置をリセットするメソッド
        public void CharOffsetReset()
        {
            kryptonRibbonGroupNumericUpDown3.Value = 0;
        }

        //すべての文字列のインデントと文字の上下位置をリセットする処理
        private void kryptonContextMenuItem38_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectAll();
            IndentReset();
            CharOffsetReset();
            richTextBox.DeselectAll();
        }

        //選択部分の文字列のインデントと文字の上下位置をリセットする処理
        private void kryptonContextMenuItem39_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            IndentReset();
            CharOffsetReset();
        }

        //すべての文字列のインデントのみをリセットする処理
        private void kryptonContextMenuItem40_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectAll();
            IndentReset();
            richTextBox.DeselectAll();
        }

        //選択部分の文字列のインデントのみをリセットする処理
        private void kryptonContextMenuItem41_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            IndentReset();
        }

        //すべての文字列の文字の上下位置のみをリセットする処理
        private void kryptonContextMenuItem42_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectAll();
            CharOffsetReset();
            richTextBox.DeselectAll();
        }

        //選択部分の文字列の文字の上下位置のみをリセットする処理
        private void kryptonContextMenuItem43_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            CharOffsetReset();
        }

        //スペルチェック機能
        private void kryptonRibbonGroupButton29_Click(object sender, EventArgs e)
        {

            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            using (Ookii.Dialogs.WinForms.ProgressDialog progressDialog = new ProgressDialog())
            {
                progressDialog.WindowTitle = "処理中...";
                progressDialog.Text = "ドキュメント全文のスペルチェックを実行中です";
                progressDialog.Description = "この処理には時間がかかります。しばらくお待ちください...";
                progressDialog.Animation = AnimationResource.GetShellAnimation(ShellAnimation.SearchFlashlight);
                progressDialog.DoWork += ProgressDialog_DoWork;
                progressDialog.ShowCancelButton = false;
                //重たい処理を実行
                progressDialog.ProgressBarStyle = Ookii.Dialogs.WinForms.ProgressBarStyle.MarqueeProgressBar;
                //ダイアログを表示
                progressDialog.Show();

                this.spellCheckerAdv1.SpellCheck(new SpellCheckerAdvEditorWrapper(richTextBox));
            }

        }


        private void ProgressDialog_DoWork(object sender, DoWorkEventArgs e)
        {
            while (this.Enabled == false)
            {
                ((ProgressDialog)sender).Dispose();
                break;
            }

        }


        //表示タブ

        //新しいウィンドウでこのアプリを起動する処理
        private void kryptonRibbonGroupButton41_Click(object sender, EventArgs e)
        {
            Process.Start(System.Windows.Forms.Application.ExecutablePath);
        }

        //子ウィンドウを重ねて表示する処理
        private void kryptonContextMenuItem44_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.Cascade);
        }

        //子ウィンドウを縦に並べて表示する処理
        private void kryptonContextMenuItem45_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileVertical);
        }

        //子ウィンドウを横に並べて表示する処理
        private void kryptonContextMenuItem46_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileHorizontal);
        }

        //アイコン表示にする処理
        private void kryptonContextMenuItem47_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.ArrangeIcons);
        }

        //フルスクリーン表示にする処理
        private void kryptonRibbonGroupButton43_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton43.Checked == true)
            {
                this.WindowState = FormWindowState.Normal;
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                this.AllowFormChrome = false;
                kryptonRibbon1.AllowFormIntegrate = false;
            }
            else
            {

                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.WindowState = FormWindowState.Normal;
                this.AllowFormChrome = true;
                if (Properties.Settings.Default.IsUseSystemTitleBar == true)
                {
                    kryptonRibbon1.AllowFormIntegrate = true;
                }
                else if (Properties.Settings.Default.IsUseSystemTitleBar == false)
                {
                    kryptonRibbon1.AllowFormIntegrate = false;
                }
            }
        }

        //コンパクトモードを解除する処理
        private void コンパクトモードを終了ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            menuStrip1.Hide();
            this.AllowFormChrome = true;
            kryptonRibbon1.Show();

            if (Properties.Settings.Default.IsUseSystemTitleBar == true)
            {
                kryptonRibbon1.AllowFormIntegrate = true;
            }
            else if (Properties.Settings.Default.IsUseSystemTitleBar == false)
            {
                kryptonRibbon1.AllowFormIntegrate = false;
            }
            this.Show();
        }

        //コンパクトモードにする処理
        private void kryptonRibbonGroupButton44_Click(object sender, EventArgs e)
        {
            this.Hide();
            menuStrip1.Show();
            this.AllowFormChrome = true;
            kryptonRibbonGroupButton43.Checked = false;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            kryptonRibbon1.Hide();
            kryptonRibbon1.AllowFormIntegrate = false;

            kryptonPanel11.Hide();
            splitter1.Visible = false;
            kryptonRibbonGroup1.DialogBoxLauncher = true;

            splitter2.Visible = false;
            kryptonNavigator1.Visible = false;


            kryptonRibbonGroupButton5.Checked = false;
            kryptonRibbonGroupButton6.Checked = false;
            kryptonRibbonGroupButton16.Checked = false;
            this.Show();

        }

        private void radialMenu1_MenuVisibilityChanging(object sender, MenuVisibilityChanging e)
        {
        }

        private void toolStripPanelItem4_Click(object sender, EventArgs e)
        {

        }

        //ミニツールバー関連のイベントハンドラ

        private void toolStripComboBox2_TextUpdate(object sender, EventArgs e)
        {
            Form activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            try
            {
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new System.Drawing.Font(toolStripComboBox2.Text, richTextBox.SelectionFont.Size, richTextBox.SelectionFont.Style);
            }
            catch
            { }

            kryptonRibbonGroupComboBox1.Text = toolStripComboBox2.Text;
        }

        //フォントサイズ用のコンボボックスが閉じたときの処理
        private void toolStripComboBox3_TextUpdate(object sender, EventArgs e)
        {
            //String型からFloat型へTryParseし選択したフォントサイズのみ変更する
            if (float.TryParse(toolStripComboBox3.Text, out float rtb_FontSize))
            {
                Form activeForm = this.ActiveMdiChild;
                richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
                richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
                richTextBox.SelectionFont = new System.Drawing.Font(richTextBox.Font.Name, rtb_FontSize, richTextBox.SelectionFont.Style);
            }

            //ツールストリップのコンボボックスにも反映させる
            kryptonRibbonGroupComboBox2.Text = toolStripComboBox3.Text;
        }



        //ミニツールバーの文字拡大ボタンの処理
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            float f = float.Parse(toolStripComboBox3.Text);
            if (f < 8 | f == 8)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 9, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "9";
            }
            else if (f < 9 | f == 9)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10";
            }
            else if (f < 10 | f == 10)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10.5f, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10.5";
            }
            if (f < 10.5f | f == 10.5f)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 11, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "11";
            }
            else if (f < 11 | f == 11)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "12";
            }
            else if (f < 12 | f == 12)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 14, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "14";
            }
            else if (f < 14 | f == 14)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 16, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "16";
            }
            else if (f < 16 | f == 16)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 18, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "18";
            }
            else if (f < 18 | f == 18)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 20, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "20";
            }
            else if (f < 20 | f == 20)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 22, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "22";
            }
            else if (f < 22 | f == 22)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 24, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "24";
            }
            else if (f < 24 | f == 24)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 26, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "26";
            }
            else if (f < 26 | f == 26)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 28, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "28";
            }
            else if (f < 28 | f == 28)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 36, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "36";
            }
            else if (f < 36 | f == 36)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 48, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "48";
            }
            else if (f < 48 | f == 48)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 72, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "72";
            }

            //リボンのフォントサイズコンボボックスにも反映させる
            kryptonRibbonGroupComboBox2.Text = toolStripComboBox3.Text;


        }

        //ミニツールバーの文字縮小ボタンの処理
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            float f = float.Parse(toolStripComboBox3.Text);
            if (f > 72 | f == 72)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 48, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "48";
            }
            else if (f > 48 | f == 48)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 36, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "36";
            }
            else if (f > 36 | f == 36)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 28, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "28";
            }
            else if (f > 28 | f == 28)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 26, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "26";
            }
            else if (f > 26 | f == 26)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 24, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "24";
            }
            else if (f > 24 | f == 24)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 22, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "22";
            }
            else if (f > 22 | f == 22)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 20, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "20";
            }
            else if (f > 20 | f == 20)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 18, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "18";
            }
            else if (f > 18 | f == 18)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 16, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "16";
            }
            else if (f > 16 | f == 16)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 14, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "14";
            }
            else if (f > 14 | f == 14)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 12, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "12";
            }
            else if (f > 12 | f == 12)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 11, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "11";
            }
            else if (f > 11 | f == 11)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10.5f, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10.5";
            }
            else if (f > 10.5f | f == 10.5f)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 10, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "10";
            }
            else if (f > 10 | f == 10)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 9, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "9";
            }
            else if (f > 9 | f == 9)
            {
                richTextBox.SelectionFont = new Font(richTextBox.SelectionFont.Name, 8, richTextBox.SelectionFont.Style);
                kryptonRibbonGroupComboBox2.Text = "8";
            }

            //リボンのフォントサイズコンボボックスにも反映させる
            kryptonRibbonGroupComboBox2.Text = toolStripComboBox3.Text;
        }


        //ミニツールバーの太字ボタンの処理
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (toolStripButton1.Checked == true)
            {
                kryptonRibbonGroupClusterButton1.Checked = true;
            }
            else
            {
                kryptonRibbonGroupClusterButton1.Checked = false;
            }
            kryptonRibbonGroupClusterButton1_Click(sender, e);
        }

        //ミニツールバーの斜体ボタンの処理
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (toolStripButton2.Checked == true)
            {
                kryptonRibbonGroupClusterButton2.Checked = true;
            }
            else
            {
                kryptonRibbonGroupClusterButton2.Checked = false;
            }
            kryptonRibbonGroupClusterButton2_Click(sender, e);
        }

        private void toolStripPanelItem1_MouseEnter(object sender, EventArgs e)
        {
            IsUsingMiniToolBar = true;
        }

        private void toolStripPanelItem1_MouseLeave(object sender, EventArgs e)
        {
            IsUsingMiniToolBar = false;
        }

        //ミニツールバーの下線ボタンの処理
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItem2.Checked == true)
            {
                kryptonContextMenuItem23.Checked = true;
            }
            else
            {

                kryptonContextMenuItem23.Checked = false;
            }
            kryptonContextMenuItem23_Click(sender, e);

        }

        //ミニツールバーの打ち消し線ボックスの処理
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItem3.Checked == true)
            {
                kryptonContextMenuItem24.Checked = true;
            }
            else
            {

                kryptonContextMenuItem24.Checked = false;
            }
            kryptonContextMenuItem24_Click(sender, e);
        }

        //ミニツールバーの文字色を変更する処理
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            richTextBox.SelectionColor = kryptonRibbonGroupClusterColorButton1.SelectedColor;

        }

        // 以下は Form1 クラス内に追加してください

        // ミニツールバー用のカラー選択コントロールとボタンのフィールド
        private ToolStripSplitButton toolStripSplitButtonColorPicker;
        private Syncfusion.Windows.Forms.Tools.ColorPickerUIAdv miniColorPicker;

        private ToolStripSplitButton toolStripSplitButtonColorPicker2;
        private Syncfusion.Windows.Forms.Tools.ColorPickerUIAdv miniColorPicker2;

        // ミニツールバーに ColorPicker を追加する初期化メソッド
        private void AddColorPickerToMiniToolBar()
        {


            // ColorPicker の作成
            miniColorPicker = new Syncfusion.Windows.Forms.Tools.ColorPickerUIAdv
            {
                BorderStyle = BorderStyle.FixedSingle,
                Size = new Size(220, 180)
            };
            miniColorPicker.Picked += MiniColorPicker_SelectedColorChanged;

            

            // ToolStrip にホストするための ToolStripControlHost
            var host = new ToolStripControlHost(miniColorPicker)
            {
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoSize = true,
                Size = miniColorPicker.Size
            };
            // ToolStripSplitButton を作成してドロップダウンにホストを追加
            toolStripButton6.DropDownItems.Add(host);


            // ColorPicker の作成
            miniColorPicker2 = new Syncfusion.Windows.Forms.Tools.ColorPickerUIAdv
            {
                BorderStyle = BorderStyle.FixedSingle,
                Size = new Size(220, 180)
            };
            miniColorPicker2.Picked += MiniColorPicker2_SelectedColorChanged;

            // ToolStrip にホストするための ToolStripControlHost
            var host2 = new ToolStripControlHost(miniColorPicker2)
            {
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoSize = true,
                Size = miniColorPicker2.Size
            };

            // ToolStripSplitButton を作成してドロップダウンにホストを追加
            toolStripButton7.DropDownItems.Add(host2);

            miniColorPicker.ThemeGroup.Name = "テーマの色";
            miniColorPicker.StandardGroup.Name = "標準の色";
            miniColorPicker.MoreColorsButton.Name = "他の色...";
            miniColorPicker.StateButton.Name = "自動";

            miniColorPicker2.ThemeGroup.Name = "テーマの色";
            miniColorPicker2.StandardGroup.Name = "標準の色";
            miniColorPicker2.MoreColorsButton.Name = "他の色...";
            miniColorPicker.StateButton.Name = "自動";
        }

        // ColorPicker で色が選択されたときのハンドラ
        private void MiniColorPicker_SelectedColorChanged(object sender, EventArgs e)
        {
            Color selected = miniColorPicker.SelectedColor;

            // 現在アクティブな MDI の RichTextBox に反映
            try
            {
                Form activeform = this.ActiveMdiChild;
                if (activeform != null)
                {
                    var rtb = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    if (rtb != null)
                    {
                        rtb.Select(rtb.SelectionStart, rtb.SelectionLength);
                        rtb.SelectionColor = selected;
                        kryptonRibbonGroupClusterColorButton1.SelectedColor = selected;
                    }
                }
            }
            catch
            {
                // 安全のため無視（必要ならログ）
            }
        }


        // ColorPicker で色が選択されたときのハンドラ
        private void MiniColorPicker2_SelectedColorChanged(object sender, EventArgs e)
        {
            Color selected = miniColorPicker2.SelectedColor;

            // 現在アクティブな MDI の RichTextBox に反映
            try
            {
                Form activeform = this.ActiveMdiChild;
                if (activeform != null)
                {
                    var rtb = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
                    if (rtb != null)
                    {
                        rtb.Select(rtb.SelectionStart, rtb.SelectionLength);
                        rtb.SelectionBackColor = selected;
                        kryptonRibbonGroupClusterColorButton2.SelectedColor = selected;
                    }
                }
            }
            catch
            {
                // 安全のため無視（必要ならログ）
            }
        }



        //ミニツールバーの蛍光ペンの色を変更する処理
        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.Select(richTextBox.SelectionStart, richTextBox.SelectionLength);
            richTextBox.SelectionBackColor = kryptonRibbonGroupClusterColorButton2.SelectedColor;

        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (toolStripButton5.Checked == true)
            {
                kryptonRibbonGroupClusterButton13.Checked = true;
            }
            else
            {

                kryptonRibbonGroupClusterButton13.Checked = false;
            }
            kryptonRibbonGroupClusterButton13_Click(sender, e);
        }

        public bool IsMiniToolBarVisible { get; set; } = false;
        private void miniToolBar1_Opening(object sender, CancelEventArgs e)
        {
            richTextBox_ScanFontStyle();
            IsMiniToolBarVisible = true;
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "MiniToolBar")
            {
                ミニツールバーに移動ToolStripMenuItem.Text = "ミニツールバーに移動...";
            }
            else if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "RadialMenu")
            {
                ミニツールバーに移動ToolStripMenuItem.Text = "ラジアルメニューに移動...";

            }

        }

        private void contextMenuStrip1_Opened(object sender, EventArgs e)
        {
        }

        //ミニツールバー表示中でもリッチテキストボックスのキー操作を受け付けるようにする
        private void toolStripPanelItem1_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void contextMenuStrip1_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void kryptonRibbonGroupButton2_Click_1(object sender, MouseEventArgs e)
        {
        }

        //リッチテキストボックスの拡大率を25%に戻す処理
        private void kryptonRibbonGroupButton40_Click(object sender, EventArgs e)
        {
            kryptonTrackBar1.Value = 5;
        }

        //リッチテキストボックスの左端の余白を表示する
        private void kryptonRibbonGroupCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (kryptonRibbonGroupCheckBox1.Checked == true)
            {
                richTextBox.ShowSelectionMargin = true;
            }
            else
            {
                richTextBox.ShowSelectionMargin = false;
            }

        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            Point cliantPoint = new Point();
            Point scr = miniToolBar1.PointToScreen(cliantPoint);
            IsUsingMiniToolBar = false;
            miniToolBar1.Hide();
            contextMenuStrip1.Show(scr);
        }

        private void ミニツールバーに移動ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Point cliantPoint = new Point();
            Point scr = contextMenuStrip1.PointToScreen(cliantPoint);
            IsUsingMiniToolBar = true;
            contextMenuStrip1.Hide();

            if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "MiniToolBar")
            {
                miniToolBar1.Show(scr);
            }
            else if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "RadialMenu")
            {
                radialMenu1.ShowRadialMenu(scr);
                radialMenu1.MenuVisibility = true;
            }
        }

        RtbZoonFactorDialog rtbZoonFactorDialog;
        private void kryptonRibbonGroupButton39_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton39.Checked == true)
            {
                rtbZoonFactorDialog = new RtbZoonFactorDialog();
                rtbZoonFactorDialog.Show();
            }
            else
            {
                rtbZoonFactorDialog.Close();
            }

        }

        //ラジアルメニュー関連のイベントハンドラ

        //ラジアルメニューの太字ボタンがクリックされたときの処理
        private void radialMenuItem1_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (radialMenuItem1.Checked == false)
            {
                kryptonRibbonGroupClusterButton1.Checked = true;
            }
            else
            {
                kryptonRibbonGroupClusterButton1.Checked = false;
            }
            kryptonRibbonGroupClusterButton1_Click(sender, e);

        }

        //ラジアルメニューの斜体ボタンがクリックされたときの処理
        private void radialMenuItem2_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (radialMenuItem2.Checked == false)
            {
                kryptonRibbonGroupClusterButton2.Checked = true;
            }
            else
            {
                kryptonRibbonGroupClusterButton2.Checked = false;
            }
            kryptonRibbonGroupClusterButton2_Click(sender, e);
        }

        //ラジアルメニューの下線ボタンがクリックされたときの処理
        private void radialMenuItem3_Click(object sender, EventArgs e)
        {
            if (radialMenuItem3.Checked == false)
            {
                kryptonContextMenuItem23.Checked = true;
            }
            else
            {
                kryptonContextMenuItem23.Checked = false;
            }
            kryptonContextMenuItem23_Click(sender, e);

        }

        //ラジアルメニューの打ち消し線ボタンがクリックされたときの処理
        private void radialMenuItem4_Click(object sender, EventArgs e)
        {
            if (radialMenuItem4.Checked == false)
            {
                kryptonContextMenuItem24.Checked = true;
            }
            else
            {
                kryptonContextMenuItem24.Checked = false;
            }
            kryptonContextMenuItem24_Click(sender, e);
        }

        //ラジアルメニューのフォント選択ボタンがクリックされたときの処理
        private async void radialFontListBox1_RadialFontChanged(object sender, SlectedIndexChangedEventArgs e)
        {
            this.Activate();
            await Task.Delay(100);
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionFont = new Font(e.SelectedFontName, richTextBox.Font.Size, richTextBox.Font.Style);
            kryptonRibbonGroupComboBox1.Text = e.SelectedFontName;
            toolStripComboBox2.Text = e.SelectedFontName;
        }


        private void radialMenu1_CloseUp(object sender, PopupClosedEventArgs e)
        {
        }



        //ラジアルメニューのフォントサイズ変更ボタンがクリックされたときの処理
        private void radialMenuSlider1_SliderValueChanged(object sender, RadialMenuSliderValueChangedEventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            string s = e.Value.ToString();
            if (float.TryParse(s, out float f))
            {
                kryptonRibbonGroupComboBox2.Text = e.Value.ToString();
                toolStripComboBox3.Text = e.Value.ToString();
                richTextBox.SelectionFont = new System.Drawing.Font(richTextBox.SelectionFont.Name, f, richTextBox.SelectionFont.Style);
            }

        }

        private void radialMenuItem15_Click_1(object sender, EventArgs e)
        {

        }

        private void radialMenuSlider1_NextMenuOpened(object sender, EventArgs e)
        {
            radialMenuSlider1.SliderValue = richTextBox.SelectionFont.Size;
        }

        //ラジアルメニューの文字色に関する処理
        private void radialColorPalette1_ColorPaletteColorChanged(object sender, SelctedColorChangedEventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionColor = e.SelectedColor;
        }

        //ラジアルメニューの蛍光ペンに関する処理
        private void radialColorPalette2_ColorPaletteColorChanged(object sender, SelctedColorChangedEventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.SelectionBackColor = e.SelectedColor;

        }

        //暗号化して保存するイベントハンドラ
        private void kryptonContextMenuItem49_Click(object sender, EventArgs e)
        {

            IsFormClosing = false;
            activeform = this.ActiveMdiChild;
            richTextBox = activeform.Controls.OfType<RichTextBox>().FirstOrDefault();
            fname = activeform.Text;

            activeform.Activate();
            string str = activeform.Text.Replace("*", "");

            // 保存処理
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Title = "保存する場所を選択...", Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt", FilterIndex = Properties.Settings.Default.DefaultSaveMode + 1 })
            {
                string str1 = activeform.Text.Replace("*", "");

                if (File.Exists(str1) == true)
                {
                    saveFileDialog.FileName = str1;
                }
                else
                {
                    if (richTextBox?.Text.Length > 0)
                    {
                        saveFileDialog.FileName = richTextBox?.Lines[0];
                    }
                    else
                    {
                        saveFileDialog.FileName = "無題";
                    }
                }

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (saveFileDialog.FilterIndex == 1)
                    {
                        richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.RichText);
                    }
                    else if (saveFileDialog.FilterIndex == 2)
                    {
                        richTextBox?.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
                    }




                    // 保存が完了したらタイトルから*を外す
                    activeform.Text = saveFileDialog.FileName;

                    //保存変更確認基準の更新
                    filepath = saveFileDialog.FileName;
                    ReadOnlyRTB.LoadFile(saveFileDialog.FileName);

                    Thread.Sleep(2000);
                    //ファイルを暗号化して保存
                    EncryptOrFallback(saveFileDialog.FileName);


                }
            }

        }

        private void EncryptOrFallback(string path)
        {
            try
            {
                // 標準のEFS暗号化を試す
                File.Encrypt(path);

                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "暗号化保存完了", InstructionText = "ドキュメントのAES暗号化保存が完了しました", Text = "これにより保存したドキュメントファイルは暗号化して保存され、現在ログイン中のユーザーのみこのファイルにアクセスできます。ログイン中のユーザー以外のアカウントがファイルにアクセスしようとするとアクセス拒否が発生する場合がありますのでご注意ください。\r\n暗号化したファイルにアクセス、保存、復号化を行えるのはこのファイルを保存したユーザーのみです。", OwnerWindowHandle = this.Handle }) { taskDialog.Show(); }
                        ;
            }
            catch (IOException)
            {
                // EFS不可なら DPAPI を使ってファイル内容を保護して別ファイルに保存する（フォールバック）
                byte[] data = File.ReadAllBytes(path);
                byte[] protectedData = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(path, protectedData);
                // 必要なら元ファイルを上書き／削除する等の処理をここで行う

                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "暗号化保存完了", InstructionText = "ドキュメントのDPAPI暗号化保存が完了しました", Text = "これにより保存したドキュメントファイルは暗号化して保存され、現在ログイン中のユーザーのみこのファイルにアクセスできます。ログイン中のユーザー以外のアカウントがファイルにアクセスしようとするとアクセス拒否が発生する場合がありますのでご注意ください。\n\n暗号化したファイルにアクセス、保存、復号化を行えるのはこのファイルを保存したユーザーのみです。", OwnerWindowHandle = this.Handle }) { taskDialog.Show(); }
                        ;
            }
            catch (Exception ex)
            {
                // 予期しない例外はユーザー表示等で通知
                MessageBox.Show("暗号化に失敗しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public string message = string.Empty;
        //ファイルの復号化
        private async void kryptonContextMenuItem50_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Multiselect = true, Title = "暗号化を解除するファイルの場所を選択...", Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt|リッチテキストファイルと書式なしテキストファイル(*.rtf;*.txt)|*.rtf;*.txt" })
            {


                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        var availableWindowsHello = await UserConsentVerifier.CheckAvailabilityAsync();
                        if (availableWindowsHello != UserConsentVerifierAvailability.Available)
                        {
                            return;
                        }
                        else
                        {
                            var result = await UserConsentVerifier.RequestVerificationAsync("ファイルの復号化を行うにはWindows Helloで本人確認を行う必要があります。認証してください。");

                            if (result == UserConsentVerificationResult.Verified)
                            {
                                //認証出来た場合
                                //ファイルの復号化処理
                                try
                                {
                                    int i = 0;

                                    while (openFileDialog.FileNames.Length > 0)
                                    {


                                        if (!File.Exists(openFileDialog.FileNames[i]))
                                        {
                                            message = "ファイルが存在しません。";
                                        }

                                        // まず EFS (File.Decrypt) を試す（EFS で保護されている場合）
                                        try
                                        {
                                            File.Decrypt(openFileDialog.FileNames[i]);
                                        }
                                        catch (UnauthorizedAccessException) { /* EFS の権限がない */ }
                                        catch (IOException) { /* EFS でない可能性 */ }
                                        catch (Exception) { /* 無視して DPAPI を試す */ }

                                        // 次に DPAPI (ProtectedData.Unprotect) を試す。元が ProtectedData.Protect(..., null, CurrentUser) の場合を想定。
                                        try
                                        {
                                            // 重要: 上書き前にバックアップを作る
                                            string backup = openFileDialog.FileNames[i] + ".bak";
                                            if (!File.Exists(backup))
                                            {
                                                File.Copy(openFileDialog.FileNames[i], backup);
                                            }

                                            byte[] protectedBytes = File.ReadAllBytes(openFileDialog.FileNames[i]);

                                            // optionalEntropy を使って保護している場合は同じ値をここに渡す必要があります（この例では null）
                                            byte[] unprotected = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);

                                            // 復号に成功したら上書き保存（必要なら別ファイル名で復元する）
                                            File.WriteAllBytes(openFileDialog.FileNames[i], unprotected);

                                            message = "DPAPI による復号に成功しました（ProtectedData.Unprotect）。";
                                        }
                                        catch (CryptographicException ex)
                                        {
                                            message = "DPAPI の復号に失敗しました。現在のユーザー／マシンでは復号できない可能性があります: " + ex.Message;
                                        }
                                        catch (Exception ex)
                                        {
                                            message = "復号処理で予期しないエラーが発生しました: " + ex.Message;
                                        }


                                        i++;
                                        if (i == openFileDialog.FileNames.Length)
                                        {
                                            //復号化が終わったら必ずメッセージを表示
                                            using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "復号化完了", InstructionText = openFileDialog.FileNames.Length + "件のファイルの復号化が完了しました", Text = "これによりどのユーザーでもこのファイルにアクセスできるようになります。また指定した各ファイルの場所にバックアップファイル(*.bak)を作成しました。", OwnerWindowHandle = this.Handle }) { taskDialog.Show(); }
                                            ;
                                            break;
                                        }
                                    }
                                }
                                catch
                                {
                                    MessageBox.Show("ファイルの復号化中にエラーが発生しました: " + message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }

                        }
                        else
                        {
                            //認証をキャンセルした場合
                            return;
                        }
                    }
                }
            }
        }

        System.Windows.Forms.Timer SaveTimer = new System.Windows.Forms.Timer();
        public void AppLoadSettings()
        {

            //ユーザーインターフェースのオプション

            //ミニツールバーとラジアルメニュー
            if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "MiniToolBar")
            {
                radialMenu1.HidePopup();
            }
            else if (Properties.Settings.Default.MiniToolBarOrRadialMenu == "RadialMenu")
            {
                miniToolBar1.Hide();
            }

            //リボンのタッチ操作向けUI調整
            if (Properties.Settings.Default.IsOptimizedTachModeRibbon == true)
            {
                kryptonRibbon1.StateCommon.RibbonGeneral.TextFont = new Font("Yu Gothic UI", 12.5F, FontStyle.Regular);
            }
            else if (Properties.Settings.Default.IsOptimizedTachModeRibbon == false)
            {
                kryptonRibbon1.StateCommon.RibbonGeneral.TextFont = new Font("Yu Gothic UI", 11.25F, FontStyle.Regular);
            }

            try
            {
                //テーマ
                //既定のテーマ
                if (Properties.Settings.Default.Theme == "Global")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.Global;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Default;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Blue;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Blue;
                }
                //Professional-Systemテーマ
                else if (Properties.Settings.Default.Theme == "ProfessionalSystem")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.ProfessionalSystem;

                    miniToolBar1.Style = ToolStripExStyle.Office2016White;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Silver;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Office2016Colorful;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2003;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2003;
                }
                //Professional-Office2003テーマ
                else if (Properties.Settings.Default.Theme == "ProfessionalOffice2003")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.ProfessionalOffice2003;

                    miniToolBar1.Style = ToolStripExStyle.Office2016White;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Silver;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Office2016Colorful;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2003;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2003;
                }
                //Office2007Blueテーマ
                else if (Properties.Settings.Default.Theme == "Office2007Blue")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.Office2007Blue;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Blue;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Blue;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Blue;
                }
                //Office2007Silverテーマ
                else if (Properties.Settings.Default.Theme == "Office2007Silver")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.Office2007Silver;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Silver;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Silver;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Silver;

                }
                //Office2007Blackテーマ
                else if (Properties.Settings.Default.Theme == "Office2007Black")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.Office2007Black;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Black;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                }
                //Office2010Blueテーマ
                else if (Properties.Settings.Default.Theme == "Office2010Blue")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.Office2010Blue;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Blue;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Blue;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Blue;
                }
                //Office2010Silverテーマ
                else if (Properties.Settings.Default.Theme == "Office2010Silver")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.Office2010Silver;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Silver;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Silver;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Silver;
                }
                //Office2010Blackテーマ
                else if (Properties.Settings.Default.Theme == "Office2010Black")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.Office2010Black;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Black;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                }
                //SparkleBlueテーマ
                else if (Properties.Settings.Default.Theme == "SparkleBlue")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.SparkleBlue;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Black;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                }
                //SparkleOrangeテーマ
                else if (Properties.Settings.Default.Theme == "SparkleOrange")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.SparkleOrange;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Black;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                }
                //SparklePurpleテーマ
                else if (Properties.Settings.Default.Theme == "SparklePurple")
                {
                    statusStripEx1.Refresh();
                    statusStripEx1.ResetOfficeColorScheme();

                    kryptonPalette1.BasePaletteMode = PaletteMode.SparklePurple;

                    miniToolBar1.Style = ToolStripExStyle.Default;
                    statusStripEx1.OfficeColorScheme = ToolStripEx.ColorScheme.Black;
                    statusStripEx1.VisualStyle = StatusStripExStyle.Default;

                    miniColorPicker.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                    miniColorPicker2.Style = (ColorPickerUIAdv.visualstyle)Style.Office2007Black;
                }

            }
            catch
            {
                //無視
                this.Refresh();
            }

            //システムスタイルのタイトルバー
            if (Properties.Settings.Default.IsUseSystemTitleBar == true)
            {
                kryptonRibbon1.AllowFormIntegrate = true;
            }
            else if (Properties.Settings.Default.IsUseSystemTitleBar == false)
            {
                kryptonRibbon1.AllowFormIntegrate = false;
            }

            //Office2007のリボンシェイプ
            if (Properties.Settings.Default.IsUseOffice2007RibbonShape == true)
            {
                kryptonRibbon1.StateCommon.RibbonGeneral.RibbonShape = PaletteRibbonShape.Office2007;
            }
            else if (Properties.Settings.Default.IsUseOffice2007RibbonShape == false)
            {
                kryptonRibbon1.StateCommon.RibbonGeneral.RibbonShape = PaletteRibbonShape.Inherit;
            }

            kryptonPalette2.BasePaletteMode = kryptonPalette1.BasePaletteMode;
            //ダイアログやMdiウィンドウのシステムスタイルのタイトルバー使用
            if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == true)
            {
                kryptonPalette2.AllowFormChrome = InheritBool.False;
            }
            else if (Properties.Settings.Default.IsUseDialogAndMdiWindowOnSystemTitleBar == false)
            {
                kryptonPalette2.AllowFormChrome = InheritBool.True;
            }



            //起動時の設定(Form1_Loadに実装済)

            //フォント設定(MdiFormCreateに移動)

            //AI機能のオプション
            //AI機能の表示
            if (Properties.Settings.Default.IsUseAIFunction == true)
            {
                kryptonRibbonTab4.Visible = true;
                aI機能ToolStripMenuItem.Visible = true;
                toolStripSeparator5.Visible = true;
            }
            else if (Properties.Settings.Default.IsUseAIFunction == false)
            {
                kryptonRibbonTab4.Visible = false;
                aI機能ToolStripMenuItem.Visible = false;
                toolStripSeparator5.Visible = false;
            }

            //その他のオプション
            //ステータスバーの表示
            if (Properties.Settings.Default.IsShowStatusBar == true)
            {
                statusStripEx1.Visible = true;
            }
            else if (Properties.Settings.Default.IsShowStatusBar == false)
            {
                statusStripEx1.Visible = false;
            }

            //既定でリボンを最小化
            if (Properties.Settings.Default.IsMinimizedRibbon == true)
            {
                kryptonRibbon1.MinimizedMode = true;
            }
            else if (Properties.Settings.Default.IsMinimizedRibbon == false)
            {
                kryptonRibbon1.MinimizedMode = false;
            }

            //保存タブ
            //ドキュメントの保存
            //既定のファイル保存形式(int型で変更、各SaveFileDialogで実装)

            //自動保存機能の有効・無効
            if (Properties.Settings.Default.IsUseAutoSave == true)
            {
                //自動保存を行う間隔(整数×60000でミリ秒単位に換算、割り算で分に換算する)

                //SaveTimerを一時停止して値を切り替え実行
                if (this.MdiChildren.Length > 0)
                {
                    SaveTimer.Stop();
                    SaveTimer.Interval = Properties.Settings.Default.AutoSaveTime;
                    SaveTimer.Tick += SaveTimer_Tick;
                    SaveTimer.Start();
                }
            }
            else if (Properties.Settings.Default.IsUseAutoSave == false)
            {
                SaveTimer.Stop();
            }




            //ドキュメントを閉じたときに開いた保存を確認せずファイルを自動的に上書き保存する
            if (Properties.Settings.Default.IsDocsCloseOnSave == true)
            {
            }
            else if (Properties.Settings.Default.IsDocsCloseOnSave == false)
            {
            }

            //ファイルの関連付け(なし)
            //APIキー(なし)
            //リセット(なし)

            //クイックアクセスツールバ
            if (Properties.Settings.Default.QAT1Visible == true)
            {
                kryptonRibbonQATButton1.Visible = true;
            }
            else if (Properties.Settings.Default.QAT1Visible == false)
            {
                kryptonRibbonQATButton1.Visible = false;
            }

            if (Properties.Settings.Default.QAT2Visible == true)
            {
                kryptonRibbonQATButton2.Visible = true;
            }
            else if (Properties.Settings.Default.QAT2Visible == false)
            {
                kryptonRibbonQATButton2.Visible = false;
            }

            if (Properties.Settings.Default.QAT3Visible == true)
            {
                kryptonRibbonQATButton3.Visible = true;
            }
            else if (Properties.Settings.Default.QAT3Visible == false)
            {
                kryptonRibbonQATButton3.Visible = false;
            }

            if (Properties.Settings.Default.QAT4Visible == true)
            {
                kryptonRibbonQATButton4.Visible = true;
            }
            else if (Properties.Settings.Default.QAT4Visible == false)
            {
                kryptonRibbonQATButton4.Visible = false;
            }

            if (Properties.Settings.Default.QAT5Visible == true)
            {
                kryptonRibbonQATButton5.Visible = true;
            }
            else if (Properties.Settings.Default.QAT5Visible == false)
            {
                kryptonRibbonQATButton5.Visible = false;
            }

            if (Properties.Settings.Default.QAT6Visible == true)
            {
                kryptonRibbonQATButton6.Visible = true;
            }
            else if (Properties.Settings.Default.QAT6Visible == false)
            {
                kryptonRibbonQATButton6.Visible = false;
            }

            if (Properties.Settings.Default.QAT7Visible == true)
            {
                kryptonRibbonQATButton7.Visible = true;
            }
            else if (Properties.Settings.Default.QAT7Visible == false)
            {
                kryptonRibbonQATButton7.Visible = false;
            }

            if (Properties.Settings.Default.QAT8Visible == true)
            {
                kryptonRibbonQATButton8.Visible = true;
            }
            else if (Properties.Settings.Default.QAT8Visible == false)
            {
                kryptonRibbonQATButton8.Visible = false;
            }
        }

        //自動上書き保存の処理
        private void SaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if(this.MdiChildren.Length != 0)
                {
                    Form form = this.ActiveMdiChild;
                    richTextBox = form.Controls.OfType<RichTextBox>().FirstOrDefault();
                    string filePath = form.Text.Replace("*","");
                    if(File.Exists(filePath) == true)
                    {
                        if (filePath.Contains(".rtf") | filePath.Contains(".txt.rtf"))
                        {
                            richTextBox?.SaveFile(filePath, RichTextBoxStreamType.RichText);
                        }
                        else if (filePath.Contains(".txt") | filePath.Contains(".rtf.txt"))
                        {
                            richTextBox?.SaveFile(filePath, RichTextBoxStreamType.PlainText);
                        }

                        // 保存が完了したらタイトルから*を外す
                        activeform.Text = filePath;
                    }

                    //保存変更確認基準の更新
                    filepath = filePath;
                    ReadOnlyRTB.LoadFile(filePath);
                }
            }
            catch {//何もしない
            }
        }

        private void kryptonContextMenuItem19_Click(object sender, EventArgs e)
        {
            using (AboutBox1 aboutBox1 = new AboutBox1())
            {
                aboutBox1.ShowDialog();
            }
        }

        private void kryptonContextMenuItem18_Click(object sender, EventArgs e)
        {
            using (ThirdPartyDialog thirdPartyDialog = new ThirdPartyDialog())
            {
                thirdPartyDialog.ShowDialog();
            }
        }

        TutorialDialog tutorialDialog;
        private void kryptonContextMenuItem15_Click(object sender, EventArgs e)
        {
            tutorialDialog = new TutorialDialog();
            tutorialDialog.Show();
            tutorialDialog.StartPosition = FormStartPosition.CenterParent;
        }

        #region 罫線を引く処理
        public void DrawLine(string Line)
        {
            kryptonRibbonGroupClusterButton17.TextLine = Line;
            //罫線を引く処理
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (Line == "─" | Line == "━" | Line == "＝" | Line == "＊" | Line == "・" | Line == "═")
            {

                while (true)
                {
                    Line = Line + Line;
                    if (Line.Length > 20)
                    {
                        richTextBox.AppendText(Line + Environment.NewLine);
                        break;
                    }
                }
            }
            else
            {
                while (true)
                {
                    Line = Line + Line;
                    if (Line.Length > 60)
                    {
                        richTextBox.AppendText(Line + Environment.NewLine);
                        break;
                    }
                }
            }



        }

        private void kryptonContextMenuItem51_Click(object sender, EventArgs e)
        {
            DrawLine("─");
        }

        private void kryptonContextMenuItem52_Click(object sender, EventArgs e)
        {
            DrawLine("━");
        }

        private void kryptonContextMenuItem53_Click(object sender, EventArgs e)
        {
            DrawLine("_");
        }

        private void kryptonContextMenuItem54_Click(object sender, EventArgs e)
        {
            DrawLine("＝");
        }

        private void kryptonContextMenuItem57_Click(object sender, EventArgs e)
        {
            DrawLine("＊");
        }

        private void kryptonContextMenuItem58_Click(object sender, EventArgs e)
        {
            DrawLine("・");
        }

        private void kryptonContextMenuItem59_Click(object sender, EventArgs e)
        {
            DrawLine("|");
        }

        private void kryptonContextMenuItem60_Click(object sender, EventArgs e)
        {
            DrawLine("/");
        }

        private void kryptonContextMenuItem62_Click(object sender, EventArgs e)
        {
            DrawLine("~");
        }

        private void kryptonContextMenuItem55_Click(object sender, EventArgs e)
        {
            DrawLine("═");
        }
        #endregion

        //ノートシール関連のイベントハンドラ
        private void kryptonContextMenuImageSelect4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Clipboard.SetImage(kryptonContextMenuImageSelect4.ImageList.Images[kryptonContextMenuImageSelect4.SelectedIndex]);
            kryptonRibbonGroupClusterButton18.ImageSmall = kryptonContextMenuImageSelect4.ImageList.Images[kryptonContextMenuImageSelect4.SelectedIndex];
            richTextBox.Paste();

            try
            {
                kryptonContextMenuImageSelect4.SelectedIndex = -1;
            }
            catch
            {
                //無視
            }

        }

        private void kryptonRibbonGroupClusterButton18_Click(object sender, EventArgs e)
        {
            Clipboard.SetImage(kryptonRibbonGroupClusterButton18.ImageSmall);
            richTextBox.Paste();
        }

        private void kryptonRibbonGroupClusterButton17_Click(object sender, EventArgs e)
        {
            DrawLine(kryptonRibbonGroupClusterButton17.TextLine);
        }

        //「開発者向け機能」タブの切り替え処理
        private void kryptonRibbonGroupButton52_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton30.Checked = false;
            if (kryptonRibbonGroupButton52.Checked == true)
            {
                kryptonRibbon1.SelectedContext = "DevToolContext";
                kryptonRibbon1.SelectedTab = kryptonRibbonTab8;

                //ドッキングナビゲーターとドッキングマネージャーの設定(複数回設定すると例外発生するので1回のみ行う)
                try
                {
                    RegisterDockingPaletteHandlers();
                    kryptonDockingManager1.ManageFloating(this);
                    kryptonDockingManager1.ManageNavigator(kryptonDockableNavigator1);
                    kryptonDockingManager1.DefaultCloseRequest = DockingCloseRequest.HidePage;

                    //ラムダ式を使ってイベントハンドラを追加する
                    //ドッキングマネージャーのページが閉じられたときの処理(ドッキングナビゲーターを表示してページを追加する、ページは消さない)
                    kryptonDockingManager1.PageCloseRequest += (s, args) =>
                    {
                        kryptonRibbonGroupButton55.Checked = true;
                        splitter3.Show();
                        kryptonDockableNavigator1.Show();
                        //ドッキングマネージャーのページが閉じられたときに、閉じられたページのUniqueNameが特定の値だったらkryptonDockableNavigator1にページを追加する
                        //「コンソール」ページ
                        if (args.UniqueName == kryptonPage7.UniqueName)
                        {
                            //ページを追加する
                            kryptonDockableNavigator1.Pages.Add(kryptonDockingManager1.Pages[0]);
                            //ドッキングナビゲーターのページを切り替える
                            kryptonDockableNavigator1.SelectedPage = kryptonDockingManager1.Pages[0];
                        }

                    };
                    //ドッキングナビゲーターの選択ページが変更されたときの処理(SelectedPageを使ってタブが1つ以上ある場合にkryptonDockableNavigator1を表示する)
                    kryptonDockableNavigator1.SelectedPageChanged += (s, args) =>
                    {
                        if (kryptonDockableNavigator1.SelectedPage != null)
                        {
                            kryptonRibbonGroupButton55.Checked = true;
                            splitter3.Show();
                            kryptonDockableNavigator1.Show();

                        }
                        else
                        {
                            kryptonRibbonGroupButton55.Checked = false;
                            kryptonDockableNavigator1.Hide();
                            splitter3.Hide();
                        }
                    };
                }
                catch
                {
                    //無視
                }


            }
            else
            {
                kryptonRibbon1.SelectedContext = string.Empty;

                //ドッキングナビゲーターを非表示にする
                kryptonDockableNavigator1.Hide();
                splitter3.Hide();
            }
        }


        //ドッキングマネージャーがタブをフローティングウィンドウ表示したときにkryptonDockableNavigator1のページが-1だったらkryptonDockableNavigator1を非表示にする処理
        private void kryptonDockableNavigator1_AfterPageDrag(object sender, PageDragEndEventArgs e)
        {
            kryptonDockableNavigator1.Show();
            if (kryptonDockableNavigator1.Pages.Count > -1)
            {
                KryptonForm form = (KryptonForm)kryptonDockableNavigator1.FindForm();
                form.Palette = kryptonPalette1;
                kryptonRibbonGroupButton55.Checked = false;
                kryptonDockableNavigator1.Hide();
                splitter3.Hide();
            }
            else
            {
                kryptonRibbonGroupButton55.Checked = true;
                splitter3.Show();
                kryptonDockableNavigator1.Show();

            }
        }

        //ドッキングナビゲーターの閉じるボタンがクリックされたときの処理(ドッキングマネージャー自体を非表示にする、タブは閉じない)
        private void kryptonDockableNavigator1_CloseAction(object sender, CloseActionEventArgs e)
        {
            kryptonRibbonGroupButton55.Checked = false;
            kryptonDockableNavigator1.Hide();
            splitter3.Hide();
        }

        //ドッキングナビゲーターの表示・非表示を切り替える処理
        private void kryptonRibbonGroupButton55_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton55.Checked == true)
            {
                splitter3.Show();
                kryptonDockableNavigator1.Show();

            }
            else if (kryptonRibbonGroupButton55.Checked == false)
            {
                kryptonDockableNavigator1.Hide();
                splitter3.Hide();
            }
        }

        // イベント登録（Form1_Load の末尾かコンストラクタの InitializeComponent 後に呼ぶ）
        private void RegisterDockingPaletteHandlers()
        {
            // ワークスペース（タブホスト）が追加されたとき
            kryptonDockingManager1.FloatingWindowAdding += KryptonDockingManager1_FloatingWindowAdded;

        }


        // FloatingWindowAdded ハンドラ（存在する場合）
        private void KryptonDockingManager1_FloatingWindowAdded(object sender, FloatingWindowEventArgs e)
        {
            try
            {
                var fw = e.FloatingWindow as ComponentFactory.Krypton.Toolkit.KryptonForm;
                if (fw != null)
                {
                    fw.Palette = kryptonPalette1;
                    fw.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
                }
                var pages = fw.Controls.OfType<ComponentFactory.Krypton.Navigator.KryptonPage>().FirstOrDefault();
                if (pages != null)
                {
                    pages.Palette = kryptonPalette1;
                    pages.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
                }
            }
            catch { }
        }

        //バッチファイルやPowerShellファイルを実行する処理
        public void RubBatchAndPowershellFile(string FilePath)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo(FilePath)
            {
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                UseShellExecute = false
            };
            Process process = Process.Start(startInfo);
            process.WaitForExit();
            //実行結果を表示
            kryptonRichTextBox1.AppendText(process.StandardOutput.ReadToEnd());
            string TempPath = Path.GetTempPath();
            //使用後は必ず一時ファイルを削除する
            if (File.Exists(TempPath + "BatchFile.bat") == true)
            {
                File.Delete(TempPath + "BatchFile.bat");
            }
            else if (File.Exists(TempPath + "PowerShellFile.ps1") == true)
            {
                File.Delete(TempPath + "PowerShellFile.ps1");
            }

            //ガベージコレクションを実行
            GC.Collect();
        }

        //バッチファイルを実行する処理
        private void kryptonRibbonGroupButton48_Click(object sender, EventArgs e)
        {
            string TempPath = Path.GetTempPath();
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string script = null;
            if (richTextBox.SelectionLength > 0)
            {
                script = richTextBox.SelectedText;

            }
            else
            {
                script = richTextBox.Text;
            }

            if (richTextBox.Text.Contains("cmd") == true | richTextBox.Text.Contains("Cmd") == true | richTextBox.Text.Contains("CMD") == true | richTextBox.Text.Contains("powershell") == true | richTextBox.Text.Contains("bash") == true | richTextBox.Text.Contains("Cmd") == true | richTextBox.Text.Contains("Powershell") == true | richTextBox.Text.Contains("PowerShell") == true | richTextBox.Text.Contains("Bash") == true)
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "実行失敗", InstructionText = "この操作を正しく実行できませんでした", Text = "この機能では動的にCMDやPowerShellに切り替え・実行することはできません。このコマンドを使用する場合、ターミナルまたはコマンドプロンプトを使用してください。", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning })
                {
                    taskDialog.Show();
                    //使用後は必ず一時ファイルを削除する
                    if (File.Exists(TempPath + "BatchFile.bat") == true)
                    {
                        File.Delete(TempPath + "BatchFile.bat");
                    }
                    else if (File.Exists(TempPath + "PowerShellFile.ps1") == true)
                    {
                        File.Delete(TempPath + "PowerShellFile.ps1");
                    }
                }
            }
            else
            {
                //テンポラリフォルダにバッチファイルを作成して実行する
                using (StreamWriter streamWriter = new StreamWriter(TempPath + "BatchFile.bat"))
                {
                    streamWriter.WriteLine("@echo off");
                    streamWriter.WriteLine(script);
                    streamWriter.Close();
                    RubBatchAndPowershellFile(TempPath + "BatchFile.bat");
                }
            }



        }

        //PowerShellファイルを実行する処理
        private void kryptonRibbonGroupButton49_Click(object sender, EventArgs e)
        {

            string TempPath = Path.GetTempPath();
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string script = null;
            if (richTextBox.SelectionLength > 0)
            {
                script = richTextBox.SelectedText;

            }
            else
            {
                script = richTextBox.Text;
            }

            if (richTextBox.Text.Contains("cmd") == true | richTextBox.Text.Contains("Cmd") == true | richTextBox.Text.Contains("CMD") == true | richTextBox.Text.Contains("powershell") == true | richTextBox.Text.Contains("bash") == true | richTextBox.Text.Contains("Cmd") == true | richTextBox.Text.Contains("Powershell") == true | richTextBox.Text.Contains("PowerShell") == true | richTextBox.Text.Contains("Bash") == true)
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "実行失敗", InstructionText = "この操作を正しく実行できませんでした", Text = "この機能では動的にCMDやPowerShellに切り替え・実行することはできません。このコマンドを使用する場合、ターミナルまたはコマンドプロンプトを使用してください。", Icon = Microsoft.WindowsAPICodePack.Dialogs.TaskDialogStandardIcon.Warning })
                {
                    taskDialog.Show();
                    //使用後は必ず一時ファイルを削除する
                    if (File.Exists(TempPath + "BatchFile.bat") == true)
                    {
                        File.Delete(TempPath + "BatchFile.bat");
                    }
                    else if (File.Exists(TempPath + "PowerShellFile.ps1") == true)
                    {
                        File.Delete(TempPath + "PowerShellFile.ps1");
                    }
                }
            }
            else
            {
                //テンポラリフォルダにバッチファイルを作成して実行する
                using (StreamWriter streamWriter = new StreamWriter(TempPath + "PowerShellFile.ps1"))
                {
                    streamWriter.WriteLine("@echo off");
                    streamWriter.WriteLine(script);
                    streamWriter.Close();
                    RubBatchAndPowershellFile(TempPath + "PowerShellFile.ps1");
                }
            }
        }

        //バッチファイルやPowerShellファイルを保存する処理
        private void kryptonRibbonGroupButton53_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "バッチファイル (*.bat)|*.bat|PowerShell スクリプト (*.ps1)|*.ps1";
                saveFileDialog.Title = "スクリプトファイルを保存する場所を選択...";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string script = richTextBox.Text;
                    using (StreamWriter streamWriter = new StreamWriter(saveFileDialog.FileName))
                    {
                        streamWriter.WriteLine("@echo off");
                        streamWriter.WriteLine(script);
                        streamWriter.Close();
                        streamWriter.Dispose();
                    }
                }
            }
        }

        //チャットの送信処理(ユーザー側のチャットメッセージを追加する処理)
        private async void kryptonButton9_Click(object sender, EventArgs e)
        {
            string reply = null;
            try
            {
                kryptonTextBox5.Enabled = false;

                richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
                richTextBox1.AppendText(kryptonTextBox5.Text + Environment.NewLine);
                richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
                //AIに問い合わせて返信を取得する処理
                if (kryptonComboBox1.SelectedIndex == 0)
                {
                    ChatClient client = new ChatClient(model: "gpt-4o", apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.OpenAIChatGPTAPIKey));

                    ChatCompletion completion = client.CompleteChat(kryptonTextBox5.Text);

                    reply = $"{completion.Content[0].Text}" + Environment.NewLine;

                }
                else if (kryptonComboBox1.SelectedIndex == 1)
                {
                    var client = new Client(apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.GoogleGeminiAPIKey));
                    var response = await client.Models.GenerateContentAsync(
                      model: "gemini-3-flash-preview", contents: kryptonTextBox5.Text
                    );
                    reply = $"{response.Candidates[0].Content.Parts[0].Text}" + Environment.NewLine;
                }
                richTextBox1.AppendText(reply);


            }
            catch (Exception ex)
            {
                richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
                richTextBox1.AppendText($"エラーが発生しました:\r\n{ex.Message}" + Environment.NewLine);
            }
            finally
            {
                kryptonTextBox5.Enabled = true;
            }

        }

        //チャットの入力欄にテキストが入力されたときに送信ボタンを有効化する処理
        private void kryptonTextBox5_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox5.Text != string.Empty)
            {
                kryptonButton9.Enabled = true;

            }
            else
            {
                kryptonButton9.Enabled = false;
            }
        }

        private void kryptonPanel1_Click(object sender, EventArgs e)
        {
        }

        //チャットのクリア
        private void buttonSpecAny13_Click(object sender, EventArgs e)
        {
            richTextBox1.ResetText();
        }

        //チャットの表示・非表示を切り替える処理
        private void kryptonRibbonGroupButton50_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton50.Checked == true)
            {
                splitter2.Visible = true;
                kryptonNavigator1.Visible = true;


                kryptonRibbonGroupButton5.Checked = true;
                kryptonRibbonGroupButton6.Checked = false;
                kryptonRibbonGroupButton16.Checked = false;
                kryptonRibbonGroupButton50.Checked = true;

                kryptonNavigator1.SelectedPage = kryptonPage8;
            }
            else
            {
                buttonSpecAny4_Click(sender, e);
            }

        }

        //チャットのコピー
        private void コピーToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        //AI用の用途別のプロンプトのテンプレートを挿入する処理
        public string InsertAIPrompt(string rtbText, int promptMode)
        {
            //校正だった場合
            if (promptMode == 1)
            {
                rtbText = $"以下の文章を校正してください。誤字脱字の修正、表現の改善、文法の修正を行い、校正後の文章のみを出力してください。\r\n{rtbText}";
            }
            //要約だった場合
            else if (promptMode == 2)
            {
                rtbText = $"以下の文章を要約してください。重要なポイントを抽出し、簡潔な要約を提供してください。\r\n{rtbText}";
            }
            //丁寧語に置き換える場合
            else if (promptMode == 3)
            {
                rtbText = $"以下の文章を丁寧語に置き換えてください。敬語表現を使用し、丁寧な文章にしてください。\r\n{rtbText}";
            }
            //選択箇所のテキストを調べる場合
            else if (promptMode == 4)
            {
                rtbText = $"以下の文章または単語についてWeb上などで調べてください。\r\n{rtbText}";
            }
            //選択箇所のテキストを日本語に翻訳する場合
            if (promptMode == 5)
            {
                rtbText = $"以下の文章を日本語に翻訳し翻訳後の文章のみ出力してください。\r\n{rtbText}";
            }

            return rtbText;
        }

        //AIチャットのプロンプトを送信して返信をダイアログに表示する処理
        public void SendAIChatDialogMode(string prompt, int promptMode)
        {
            string reply = null;

            try
            {
                if (Properties.Settings.Default.OpenAIChatGPTAPIKey != null)
                {
                    ChatClient client = new ChatClient(model: "gpt-4o", apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.OpenAIChatGPTAPIKey));
                    ChatCompletion completion = client.CompleteChat(prompt);
                    reply = $"ChatGPT:{completion.Content[0].Text}" + Environment.NewLine;
                }
                else if (Properties.Settings.Default.GoogleGeminiAPIKey != null)
                {
                    var client = new Client(apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.GoogleGeminiAPIKey));
                    var response = client.Models.GenerateContentAsync(
                      model: "gemini-3-flash-preview", contents: prompt
                    ).Result;
                    reply = $"Gemini:{response.Candidates[0].Content.Parts[0].Text}" + Environment.NewLine;
                }
                else if (Properties.Settings.Default.OpenAIChatGPTAPIKey != null && Properties.Settings.Default.GoogleGeminiAPIKey != null)
                {
                    ChatClient client1 = new ChatClient(model: "gpt-4o", apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.OpenAIChatGPTAPIKey));
                    ChatCompletion completion = client1.CompleteChat(prompt);
                    reply = $"ChatGPT: {completion.Content[0].Text}" + Environment.NewLine;

                    var client2 = new Client(apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.GoogleGeminiAPIKey));
                    var response = client2.Models.GenerateContentAsync(
                      model: "gemini-3-flash-preview", contents: prompt
                    ).Result;
                    reply += Environment.NewLine + $"Gemini: {response.Candidates[0].Content.Parts[0].Text}" + Environment.NewLine;
                }
            }
            catch (Exception ex)
            {
                reply = $"エラーが発生しました:\r\n{ex.Message}" + Environment.NewLine;
            }

            string s = null;
            //校正だった場合
            if (promptMode == 1)
            {
                s = "(校正)";
            }
            //要約だった場合
            else if (promptMode == 2)
            {
                s = "(要約)";
            }
            //丁寧語に置き換える場合
            else if (promptMode == 3)
            {
                s = "(丁寧語)";
            }
            //調査だった場合
            else if (promptMode == 4)
            {
                s = "(調査)";
            }
            //翻訳だった場合
            else if (promptMode == 5)
            {
                s = "(翻訳)";
            }
            else if (promptMode == null)
            {
                s = string.Empty;
            }

            string ViewPrompt = Environment.NewLine+"送信したプロンプト:" + Environment.NewLine+prompt;


            //AIの返信をダイアログに表示する処理

            //ダイアログが存在しない場合は新しく作成して返信を表示する、存在する場合は既存のダイアログに返信を追加して表示する
            if (AIPromptAnswarDialog.Instance == null)
            {
                AIPromptAnswarDialog aIPromptAnswarDialog = new AIPromptAnswarDialog();
                aIPromptAnswarDialog.Show();
                aIPromptAnswarDialog.kryptonPage1.Text = (DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + s);
                aIPromptAnswarDialog.kryptonRichTextBox1.Text = reply+Environment.NewLine +ViewPrompt+Environment.NewLine;
            }
            else
            {
                KryptonPage page = new KryptonPage();
                page.Text = (DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + s);
                AIPromptAnswarDialog.Instance.kryptonNavigator1.Pages.Add(page);

                ComponentFactory.Krypton.Toolkit.KryptonRichTextBox kryptonRichTextBox = new ComponentFactory.Krypton.Toolkit.KryptonRichTextBox();
                kryptonRichTextBox.Dock = DockStyle.Fill;
                kryptonRichTextBox.ReadOnly = true;
                kryptonRichTextBox.Text = reply+ Environment.NewLine + ViewPrompt + Environment.NewLine;
                kryptonRichTextBox.ScrollBars = RichTextBoxScrollBars.ForcedBoth;

                kryptonRichTextBox.KryptonContextMenu = AIPromptAnswarDialog.Instance.kryptonContextMenu1;

                page.Controls.Add(kryptonRichTextBox);

                AIPromptAnswarDialog.Instance.kryptonNavigator1.SelectedPage = page;

                //ダイアログを表示する
                AIPromptAnswarDialog.Instance.Activate();
            }

        }

        //選択したテキストを校正・要約・丁寧語に置き換え直接置換する処理
        public void ReplaceSelectedTextWithAIResponse(string prompt, int promptMode)
        {
            string reply = null;
            try
            {
                if (Properties.Settings.Default.OpenAIChatGPTAPIKey != null)
                {
                    ChatClient client = new ChatClient(model: "gpt-4o", apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.OpenAIChatGPTAPIKey));
                    ChatCompletion completion = client.CompleteChat(prompt);
                    reply = $"{completion.Content[0].Text}";
                }
                else if (Properties.Settings.Default.GoogleGeminiAPIKey != null)
                {
                    var client = new Client(apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.GoogleGeminiAPIKey));
                    var response = client.Models.GenerateContentAsync(
                      model: "gemini-3-flash-preview", contents: prompt
                    ).Result;
                    reply = $"{response.Candidates[0].Content.Parts[0].Text}";
                }
                else if (Properties.Settings.Default.OpenAIChatGPTAPIKey != null && Properties.Settings.Default.GoogleGeminiAPIKey != null)
                {
                    //基本的にChatGPTを優先して使う
                    try
                    {
                        ChatClient client1 = new ChatClient(model: "gpt-4o", apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.OpenAIChatGPTAPIKey));
                        ChatCompletion completion = client1.CompleteChat(prompt);
                        reply = $"{completion.Content[0].Text}" + Environment.NewLine;
                    }
                    //何らかのエラーで使えない場合Geminiを使い、フォールバックする
                    catch 
                    {
                        var client2 = new Client(apiKey: Environment.GetEnvironmentVariable(Properties.Settings.Default.GoogleGeminiAPIKey));
                        var response = client2.Models.GenerateContentAsync(
                          model: "gemini-3-flash-preview", contents: prompt
                        ).Result;
                        reply += Environment.NewLine + $"{response.Candidates[0].Content.Parts[0].Text}" + Environment.NewLine;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBoxAdv.Show($"エラーが発生しました:\r\n{ex.Message}", "エラー", MessageBoxButtons.OK);
                //エラーが発生した場合は処理を中断する
                return;
            }
            //選択したテキストを返信で置き換える処理
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            int selectionStart = richTextBox.SelectionStart;
            int selectionLength = richTextBox.SelectionLength;
            richTextBox.Text = richTextBox.Text.Remove(selectionStart, selectionLength).Insert(selectionStart, reply);
        }

        //選択したテキストを校正する処理
        private void 校正ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            //選択したテキストを取得
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string Text = InsertAIPrompt(Environment.NewLine + richTextBox.SelectedText, 1);

            //文字が選択されていない場合は処理を中断する
            if (richTextBox.SelectionLength > 0)
            {

                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "置換方法の確認", InstructionText = "選択したテキストを校正し直接置換するかまたはダイアログ表示で校正結果を表示でしますか?", OwnerWindowHandle = this.Handle })
                {
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Replace", "直接置換");
                    taskDialog.Controls.Add(taskDialogButton);
                    //直接置換の処理
                    taskDialogButton.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                        ReplaceSelectedTextWithAIResponse(Text, 1);
                    };
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton2 = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Dialog", "ダイアログ表示");
                    taskDialog.Controls.Add(taskDialogButton2);
                    //ダイアログ表示の処理
                    taskDialogButton2.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                        SendAIChatDialogMode(Text, 1);
                    };
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton3 = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Cancel", "校正をキャンセル");
                    taskDialog.Controls.Add(taskDialogButton3);
                    //キャンセルの処理
                    taskDialogButton3.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                    };
                    taskDialog.Show();
                }

            }
            else
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "テキストが選択されていません", Text = "校正するテキストを選択してください。", OwnerWindowHandle = this.Handle})
                {
                    taskDialog.Show();
                }
            }
                


           
        }

        //選択したテキストを要約する処理
        private void 選択箇所を要約ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //選択したテキストを取得
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string Text = InsertAIPrompt(Environment.NewLine + richTextBox.SelectedText, 2);

            //文字が選択されていない場合は処理を中断する
            if (richTextBox.SelectionLength > 0)
            {
                SendAIChatDialogMode(Text, 2);
            }
            else
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "テキストが選択されていません", Text = "要約するテキストを選択してください。", OwnerWindowHandle = this.Handle })
                {
                    taskDialog.Show();
                }
            }
        }

        private void 選択箇所を丁寧語に置き換えるToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //選択したテキストを取得
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string Text = InsertAIPrompt(Environment.NewLine + richTextBox.SelectedText, 3);

            //文字が選択されていない場合は処理を中断する
            if (richTextBox.SelectionLength > 0)
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "置換方法の確認", InstructionText = "選択したテキストを丁寧語に置き換え直接置換するかまたはダイアログ表示で変換結果を表示でしますか?", OwnerWindowHandle = this.Handle })
                {
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Replace", "直接置換");
                    taskDialog.Controls.Add(taskDialogButton);
                    //直接置換の処理
                    taskDialogButton.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                        ReplaceSelectedTextWithAIResponse(Text, 3);
                    };
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton2 = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Dialog", "ダイアログ表示");
                    taskDialog.Controls.Add(taskDialogButton2);
                    //ダイアログ表示の処理
                    taskDialogButton2.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                        SendAIChatDialogMode(Text, 3);
                    };
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton3 = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Cancel", "変換をキャンセル");
                    taskDialog.Controls.Add(taskDialogButton3);
                    //キャンセルの処理
                    taskDialogButton3.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                    };
                    taskDialog.Show();
                }

            }
            else
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "テキストが選択されていません", Text = "丁寧語に変換するテキストを選択してください。", OwnerWindowHandle = this.Handle })
                {
                    taskDialog.Show();
                }
            }


        }

        //選択した画像をテキストで説明する処理
        private void 選択箇所の画像説明ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //クリップボード経由で選択した画像を取得
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            try
            {
                System.Drawing.Image img = Clipboard.GetImage();
                Bitmap bitmap = new Bitmap(img);
            }
            catch
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "画像が選択されていません", Text = "説明する画像を選択してください。", OwnerWindowHandle = this.Handle })
                {
                    taskDialog.Show();
                }
                return;
            }
        }

        //開発者向け機能タブを閉じるイベント
        private void kryptonRibbonGroupButton54_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton52.Checked = false;
            kryptonRibbon1.SelectedContext = string.Empty;

            kryptonRibbonGroupButton55.Checked = false;
            kryptonDockableNavigator1.Hide();
            splitter3.Hide();
        }

        //コンテキストメニューからAIチャットパネルを表示するイベント
        private void aIチャットパネルを表示ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splitter2.Visible = true;
            kryptonNavigator1.Visible = true;


            kryptonRibbonGroupButton5.Checked = true;
            kryptonRibbonGroupButton6.Checked = false;
            kryptonRibbonGroupButton16.Checked = false;
            kryptonRibbonGroupButton50.Checked = true;

            kryptonNavigator1.SelectedPage = kryptonPage8;
        }

        #region 印刷用のWin32API関数など
        RichTextBox printText;
        int charFrom, charTo;

        private void Pd_BeginPrint(object sender, PrintEventArgs e)
        {
            charFrom = charTo = 0;
        }

        private void Pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();

            charFrom = PrintRichText(richTextBox, charFrom, richTextBox.TextLength, e);

            if (charFrom < richTextBox.TextLength)
            {
                e.HasMorePages = true;
            }
            else
            {
                e.HasMorePages = false;
            }
        }

        public int PrintRichText(RichTextBox rtb, int charFrom, int charTo, PrintPageEventArgs e)
        {
            // RenderStruct（Win32）
            FORMATRANGE fr = new FORMATRANGE();

            // デバイスコンテキスト（描画先）
            fr.hdc = e.Graphics.GetHdc();
            fr.hdcTarget = fr.hdc;

            // 印刷可能範囲
            fr.rc = new RECT
            {
                Top = HundredthInchToTwips(e.MarginBounds.Top),
                Bottom = HundredthInchToTwips(e.MarginBounds.Bottom),
                Left = HundredthInchToTwips(e.MarginBounds.Left),
                Right = HundredthInchToTwips(e.MarginBounds.Right)
            };

            // ページ全体
            fr.rcPage = new RECT
            {
                Top = HundredthInchToTwips(e.PageBounds.Top),
                Bottom = HundredthInchToTwips(e.PageBounds.Bottom),
                Left = HundredthInchToTwips(e.PageBounds.Left),
                Right = HundredthInchToTwips(e.PageBounds.Right)
            };

            fr.chrg.cpMin = charFrom;
            fr.chrg.cpMax = charTo;

            // Win32 API 呼び出し
            IntPtr wparam = new IntPtr(1);
            IntPtr lparam = Marshal.AllocCoTaskMem(Marshal.SizeOf(fr));
            Marshal.StructureToPtr(fr, lparam, false);
            int res = SendMessage(rtb.Handle, EM_FORMATRANGE, wparam, lparam);

            // 解放
            Marshal.FreeCoTaskMem(lparam);
            e.Graphics.ReleaseHdc(fr.hdc);

            return res;
        }

        [DllImport("USER32.dll", CharSet = CharSet.Auto)]
        private static extern int SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int EM_FORMATRANGE = 0x400 + 57;

        private int HundredthInchToTwips(int n)
        {
            return (int)(n * 14.4);  // 1 inch = 1440 twips
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CHARRANGE
        {
            public int cpMin, cpMax;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FORMATRANGE
        {
            public IntPtr hdc;       // Device context
            public IntPtr hdcTarget; // Measuring device context
            public RECT rc;
            public RECT rcPage;
            public CHARRANGE chrg;
        }

        #endregion

        private void kryptonRibbonGroupButton59_Click(object sender, EventArgs e)
        {
        }

        private void kryptonRibbonGroupButton59_Click_1(object sender, EventArgs e)
        {
        }

        //選択箇所のテキストを調べる機能
        private void kryptonRibbonGroupButton59_Click_2(object sender, EventArgs e)
        {
            //選択したテキストを取得
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string Text = InsertAIPrompt(Environment.NewLine + richTextBox.SelectedText, 4);

            //文字が選択されていない場合は処理を中断する
            if (richTextBox.SelectionLength > 0)
            {
                SendAIChatDialogMode(Text, 4);
            }
            else
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "テキストが選択されていません", Text = "説明するテキストを選択してください。", OwnerWindowHandle = this.Handle })
                {
                    taskDialog.Show();
                }
            }
        }

        private void toolStripButton6_DropDownOpened(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupButton21_DropDown(object sender, ContextMenuArgs e)
        {

        }

        //編集中のドキュメントファイルの場所をエクスプローラーで開く処理
        private void kryptonContextMenuItem21_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            string s = activeForm.Text.Replace("*","");
            if(File.Exists(s) == true)
            {
                string s1 = Path.GetDirectoryName(s);
                Process.Start("explorer.exe", s1);
            }
        }

        private void buttonSpecAppMenu2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        ShortcutKeyDialog shortcutKeyDialog;
        //キーボードショートカットの一覧を表示するダイアログの表示処理
        private void buttonSpecAppMenu1_Click(object sender, EventArgs e)
        {
            buttonSpecAppMenu1.Enabled = ButtonEnabled.False;
            shortcutKeyDialog = new ShortcutKeyDialog();
            shortcutKeyDialog.Show();
           
        }

        //ToDoボタンクリックされたときの処理
        private void kryptonContextMenuItem56_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・ToDo\r\n1.\r\n2.\r\n3.\r\n");
            kryptonRibbonGroupButton25.TextLine1 = "・ToDo";
        }

        //やることリストボタンクリックされたときの処理
        private void kryptonContextMenuItem63_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・やることリスト\r\n1.\r\n2.\r\n3.\r\n");
            kryptonRibbonGroupButton25.TextLine1 = "・やることリスト";
        }

        //概要ボタンがクリックされたときの処理
        private void kryptonContextMenuItem64_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・概要");
            kryptonRibbonGroupButton25.TextLine1 = "・概要";
        }

        //要点ボタンがクリックされたときの処理
        private void kryptonContextMenuItem65_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・要点");
            kryptonRibbonGroupButton25.TextLine1 = "・要点";
        }

        //注意ボタンがクリックされたときの処理
        private void kryptonContextMenuItem66_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・注意");
            kryptonRibbonGroupButton25.TextLine1 = "・注意";
        }

        //最高ボタンがクリックされたときの処理
        private void kryptonContextMenuItem67_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・最高");
            kryptonRibbonGroupButton26.TextLine1 = "・注意";
        }

        //高ボタンがクリックされたときの処理
        private void kryptonContextMenuItem68_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・高");
            kryptonRibbonGroupButton26.TextLine1 = "・高";
        }

        //中ボタンがクリックされたときの処理
        private void kryptonContextMenuItem69_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・中");
            kryptonRibbonGroupButton26.TextLine1 = "・中";
        }

        //小ボタンがクリックされたときの処理
        private void kryptonContextMenuItem70_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・小");
            kryptonRibbonGroupButton26.TextLine1 = "・小";
        }

        //緊急ボタンがクリックされたときの処理
        private void kryptonContextMenuItem71_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・緊急");
            kryptonRibbonGroupButton26.TextLine1 = "・緊急";
        }

        //要確認ボタンがクリックされたときの処理
        private void kryptonContextMenuItem72_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・要確認");
            kryptonRibbonGroupButton26.TextLine1 = "・要確認";
        }

        //状態ボタンがクリックされたときの処理
        private void kryptonContextMenuItem73_Click(object sender, EventArgs e)
        {
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            richTextBox.AppendText("・状態");
            kryptonRibbonGroupButton26.TextLine1 = "・状態";
        }

        private void kryptonRibbonGroupButton25_Click(object sender, EventArgs e)
        {
            if(kryptonRibbonGroupButton25.TextLine1 == "・ToDo")
            {
                kryptonContextMenuItem56_Click(sender, e);
            }
            else if(kryptonRibbonGroupButton25.TextLine1 == "・やることリスト")
            {
                kryptonContextMenuItem63_Click(sender, e);
            }
            else if(kryptonRibbonGroupButton25.TextLine1 == "・概要")
            {
                kryptonContextMenuItem64_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton25.TextLine1 == "・要点")
            {
                kryptonContextMenuItem65_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton25.TextLine1 == "・注意")
            {
                kryptonContextMenuItem66_Click(sender, e);
            }
        }

        private void kryptonRibbonGroupButton26_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton26.TextLine1 == "・最高")
            {
                kryptonContextMenuItem67_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton26.TextLine1 == "・高")
            {
                kryptonContextMenuItem68_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton26.TextLine1 == "・中")
            {
                kryptonContextMenuItem69_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton26.TextLine1 == "・小")
            {
                kryptonContextMenuItem70_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton26.TextLine1 == "・緊急")
            {
                kryptonContextMenuItem71_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton26.TextLine1 == "・要確認")
            {
                kryptonContextMenuItem72_Click(sender, e);
            }
            else if (kryptonRibbonGroupButton26.TextLine1 == "・状態")
            {
                kryptonContextMenuItem73_Click(sender, e);
            }
        }

        //選択箇所のテキストを日本語に翻訳し置換する機能
        private void kryptonRibbonGroupButton60_Click(object sender, EventArgs e)
        {

            //選択したテキストを取得
            Form activeForm = this.ActiveMdiChild;
            richTextBox = activeForm.Controls.OfType<RichTextBox>().FirstOrDefault();
            string Text = InsertAIPrompt(Environment.NewLine + richTextBox.SelectedText, 5);

            //文字が選択されていない場合は処理を中断する
            if (richTextBox.SelectionLength > 0)
            {

                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "置換方法の確認", InstructionText = "選択したテキストを日本語に翻訳し直接置換するかまたはダイアログ表示で翻訳結果を表示でしますか?", OwnerWindowHandle = this.Handle })
                {
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Replace", "直接置換");
                    taskDialog.Controls.Add(taskDialogButton);
                    //直接置換の処理
                    taskDialogButton.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                        ReplaceSelectedTextWithAIResponse(Text, 5);
                    };
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton2 = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Dialog", "ダイアログ表示");
                    taskDialog.Controls.Add(taskDialogButton2);
                    //ダイアログ表示の処理
                    taskDialogButton2.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                        SendAIChatDialogMode(Text, 5);
                    };
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton taskDialogButton3 = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialogButton("Cancel", "翻訳をキャンセル");
                    taskDialog.Controls.Add(taskDialogButton3);
                    //キャンセルの処理
                    taskDialogButton3.Click += (s, arge) =>
                    {
                        taskDialog.Close();
                    };
                    taskDialog.Show();
                }

            }
            else
            {
                using (Microsoft.WindowsAPICodePack.Dialogs.TaskDialog taskDialog = new Microsoft.WindowsAPICodePack.Dialogs.TaskDialog() { Caption = "エラー", InstructionText = "テキストが選択されていません", Text = "日本語に翻訳するテキストを選択してください。", OwnerWindowHandle = this.Handle })
                {
                    taskDialog.Show();
                }
            }


        }

    }
}
