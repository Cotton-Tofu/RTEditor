using ComponentFactory.Krypton.Toolkit;
using Ookii.Dialogs.WinForms;
using SixLabors.ImageSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZXing;
using ZXing.Presentation;
using ZXing.QrCode;

namespace RTEditor
{
    public partial class BarcodeAndQRCodeScanerDialog : KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);

        public BarcodeAndQRCodeScanerDialog()
        {
            InitializeComponent();
            this.CenterToParent();

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

        string filePath = null;
        private async void kryptonButton1_Click(object sender, EventArgs a)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "PNGファイル(*.png)|*.png|JPGファイル(*.jpg)|*.jpg|EXIF情報付きJPGファイル(*.jpg)|*.jpg|JPEGファイル(*.jpeg)|*.jpeg|GIFファイル(*.gif)|*.gif|ビットマップファイル(*.bmp)|*.bmp|TIFファイル(*.tif)|*.tif|TIFFファイル(*.tiff)|*.tiff|Windows 用アイコンファイル(*.ico)|*.ico|Windows メタファイル(*.wmf)|*.wmf|Windows 拡張メタファイル(*.emf)|*.emf|WEBPファイル(*.webp)|*.webp|TGAファイル(*.tga)|*.tga|PBMファイル(*.pbm)|*.pbm|すべてのファイル(*.*)|*.*", Title = "読み込む画像ファイルを選択...", Multiselect = false })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        await Task.Run(() =>
                        {
                            kryptonRichTextBox1.ResetText();
                            Bitmap bitmap = new Bitmap(openFileDialog.FileName);
                            ZXing.BarcodeReader barcodeReader = new ZXing.BarcodeReader() { AutoRotate = true};
                            var result = barcodeReader.Decode(bitmap);
                            pictureBox1.Image = bitmap;
                            kryptonRichTextBox1.Text = result.Text;
                            filePath = openFileDialog.FileName;

                            kryptonButton9.Enabled = true;
                            kryptonButton11.Enabled = true;

                            kryptonButton4.Enabled = true;
                            kryptonButton5.Enabled = true;
                            kryptonButton8.Enabled = true;
                            kryptonButton7.Enabled = true;

                            kryptonButton2.Enabled = true;
                            kryptonButton6.Enabled = true;
                            kryptonButton3.Enabled = true;

                            this.Text = "QRコード・バーコード読み取り - "+openFileDialog.FileName;

                        });

                    }
                    catch
                    {
                    }

                }
            }
        }


        private void kryptonButton10_Click(object sender, EventArgs e)
        {
            try
            {
                ZXing.BarcodeWriter barcodeWriter = new ZXing.BarcodeWriter()
                {
                    Format = BarcodeFormat.QR_CODE,
                    Options = new QrCodeEncodingOptions
                    {
                        Height = 1000,
                        Width = 1000,
                    }
                };
                pictureBox1.Image = barcodeWriter.Write(kryptonRichTextBox1.Text);

                if (filePath != null)
                {
                    kryptonButton9.Enabled = true;
                }
                else
                {
                    kryptonButton9.Enabled = false;
                }

                kryptonButton11.Enabled = true;

                kryptonButton4.Enabled = true;
                kryptonButton5.Enabled = true;
                kryptonButton8.Enabled = true;
                kryptonButton7.Enabled = true;

                kryptonButton2.Enabled = true;
                kryptonButton6.Enabled = true;
                kryptonButton3.Enabled = true;
            }
            catch
            {

                if (kryptonRichTextBox1.Text != string.Empty)
                {
                    kryptonButton9.Enabled = false;
                    kryptonButton11.Enabled = true;

                    kryptonButton4.Enabled = true;
                    kryptonButton5.Enabled = true;
                    kryptonButton8.Enabled = true;
                    kryptonButton7.Enabled = true;

                    kryptonButton2.Enabled = true;
                    kryptonButton6.Enabled = true;
                    kryptonButton3.Enabled = true;
                }
                else
                {
                    kryptonButton9.Enabled = false;
                    kryptonButton11.Enabled = false;

                    kryptonButton4.Enabled = false;
                    kryptonButton5.Enabled = false;
                    kryptonButton8.Enabled = false;
                    kryptonButton7.Enabled = false;

                    kryptonButton2.Enabled = false;
                    kryptonButton6.Enabled = false;
                    kryptonButton3.Enabled = false;
                }

            }

            
        }

        private void kryptonButton11_Click(object sender, EventArgs e)
        {
            AllReset();
        }

        public void AllReset()
        {
            this.Text = "QRコード・バーコード読み取り";
            kryptonButton9.Enabled = false;
            kryptonButton11.Enabled = false;

            kryptonButton4.Enabled = false;
            kryptonButton5.Enabled = false;
            kryptonButton8.Enabled = false;
            kryptonButton7.Enabled = false;

            kryptonButton2.Enabled = false;
            kryptonButton6.Enabled = false;
            kryptonButton3.Enabled = false;

            pictureBox1.Image = null;
            kryptonRichTextBox1.ResetText();

            filePath = null;
        }

        private async void kryptonButton9_Click(object sender, EventArgs e)
        {

            try
            {
                await Task.Run(() =>
                {
                    kryptonRichTextBox1.ResetText();
                    Bitmap bitmap = new Bitmap(filePath);
                    ZXing.BarcodeReader barcodeReader = new ZXing.BarcodeReader() { AutoRotate = true };
                    var result = barcodeReader.Decode(bitmap);
                    pictureBox1.Image = bitmap;
                    kryptonRichTextBox1.Text = result.Text;

                    kryptonButton9.Enabled = true;
                    kryptonButton11.Enabled = true;

                    kryptonButton4.Enabled = true;
                    kryptonButton5.Enabled = true;
                    kryptonButton8.Enabled = true;
                    kryptonButton7.Enabled = true;

                    kryptonButton2.Enabled = true;
                    kryptonButton6.Enabled = true;
                    kryptonButton3.Enabled = true;

                });

            }
            catch
            {
                ResetText();
            }
        }

        private void kryptonButton12_Click(object sender, EventArgs e)
        {
            try
            {
                ZXing.BarcodeWriter barcodeWriter = new ZXing.BarcodeWriter()
                {
                    Format = BarcodeFormat.EAN_13,
                    Options = new QrCodeEncodingOptions
                    {
                        Height = 50,
                        Width =  130,
                    }
                };
                pictureBox1.Image = barcodeWriter.Write(kryptonRichTextBox1.Text);

                if (filePath != null)
                {
                    kryptonButton9.Enabled = true;
                }
                else
                {
                    kryptonButton9.Enabled = false;
                }

                kryptonButton11.Enabled = true;

                kryptonButton4.Enabled = true;
                kryptonButton5.Enabled = true;
                kryptonButton8.Enabled = true;
                kryptonButton7.Enabled = true;

                kryptonButton2.Enabled = true;
                kryptonButton6.Enabled = true;
                kryptonButton3.Enabled = true;
            }
            catch
            {
                if (kryptonRichTextBox1.Text != string.Empty)
                {
                    kryptonButton9.Enabled = false;
                    kryptonButton11.Enabled = true;

                    kryptonButton4.Enabled = true;
                    kryptonButton5.Enabled = true;
                    kryptonButton8.Enabled = true;
                    kryptonButton7.Enabled = true;

                    kryptonButton2.Enabled = true;
                    kryptonButton6.Enabled = true;
                    kryptonButton3.Enabled = true;
                }
                else
                {
                    kryptonButton9.Enabled = false;
                    kryptonButton11.Enabled = false;

                    kryptonButton4.Enabled = false;
                    kryptonButton5.Enabled = false;
                    kryptonButton8.Enabled = false;
                    kryptonButton7.Enabled = false;

                    kryptonButton2.Enabled = false;
                    kryptonButton6.Enabled = false;
                    kryptonButton3.Enabled = false;
                }
            }
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.Focus();
            kryptonRichTextBox1.Copy();
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.Focus();
            kryptonRichTextBox1.SelectAll();
            kryptonRichTextBox1.ScrollToCaret();
        }

        private void kryptonRichTextBox1_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            Process.Start(e.LinkText);
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

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt", Title = "保存する場所を選択..." })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    kryptonRichTextBox1.SaveFile(saveFileDialog.FileName);
            }
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "PNGファイル(*.png)|*.png|JPGファイル(*.jpg)|*.jpg|EXIF情報付きJPGファイル(*.jpg)|*.jpg|JPEGファイル(*.jpeg)|*.jpeg|GIFファイル(*.gif)|*.gif|ビットマップファイル(*.bmp)|*.bmp|TIFファイル(*.tif)|*.tif|TIFFファイル(*.tiff)|*.tiff|Windows 用アイコンファイル(*.ico)|*.ico|Windows メタファイル(*.wmf)|*.wmf|Windows 拡張メタファイル(*.emf)|*.emf|WEBPファイル(*.webp)|*.webp|TGAファイル(*.tga)|*.tga|PBMファイル(*.pbm)|*.pbm|すべてのファイル(*.*)|*.*", Title = "画像ファイルを保存する場所を選択...", FileName = filePath })
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    using (Bitmap bitmap = new Bitmap(pictureBox1.Image))
                    {
                        //PNGファイル
                        if (saveFileDialog.FilterIndex == 1)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Png);
                        //JPGファイル
                        else if (saveFileDialog.FilterIndex == 2)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                        //EXIF付きJPGファイル
                        else if (saveFileDialog.FilterIndex == 2)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Exif);
                        //JPEGファイル
                        else if (saveFileDialog.FilterIndex == 4)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                        //GIFファイル
                        else if (saveFileDialog.FilterIndex == 5)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Gif);
                        //ビットマップファイル
                        else if (saveFileDialog.FilterIndex == 6)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Bmp);
                        //TIFファイル
                        else if (saveFileDialog.FilterIndex == 7)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Tiff);
                        //TIFFファイル
                        else if (saveFileDialog.FilterIndex == 8)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Tiff);
                        //Windows 用アイコンファイル
                        else if (saveFileDialog.FilterIndex == 9)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Icon);
                        //Windows メタファイル
                        else if (saveFileDialog.FilterIndex == 10)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Wmf);
                        //Windows 拡張メタファイル
                        else if (saveFileDialog.FilterIndex == 11)
                            bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Emf);
                        //WEBPファイル
                        else if (saveFileDialog.FilterIndex == 12)
                        {
                            using (SixLabors.ImageSharp.Image image = SixLabors.ImageSharp.Image.Load(filePath))
                            {
                                image.SaveAsWebp(saveFileDialog.FileName);
                            }
                        }
                        //TGAファイル
                        else if (saveFileDialog.FilterIndex == 13)
                        {
                            using (SixLabors.ImageSharp.Image image = SixLabors.ImageSharp.Image.Load(filePath))
                            {
                                image.SaveAsTga(saveFileDialog.FileName);
                            }
                        }
                        //PBMファイル
                        else if (saveFileDialog.FilterIndex == 14)
                        {
                            using (SixLabors.ImageSharp.Image image = SixLabors.ImageSharp.Image.Load(filePath))
                            {
                                image.SaveAsPbm(saveFileDialog.FileName);
                            }
                        }
                        //PBMファイル
                        else if (saveFileDialog.FilterIndex == 14)
                        {
                            using (SixLabors.ImageSharp.Image image = SixLabors.ImageSharp.Image.Load(filePath))
                            {
                                image.SaveAsPbm(saveFileDialog.FileName);
                            }
                        }
                        //その他のファイル
                        else if (saveFileDialog.FilterIndex == 15)
                        {
                            using (SixLabors.ImageSharp.Image image = SixLabors.ImageSharp.Image.Load(filePath))
                            {
                                image.Save(saveFileDialog.FileName);
                            }
                        }


                    }
            }
        }

        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb == null) return;

            if (kryptonRichTextBox1.SelectedText == string.Empty)
            {
                //すべて貼り付け
                rtb.AppendText(kryptonRichTextBox1.Text);
            }
            else
            {
                //選択箇所のみ貼り付け
                rtb.AppendText(kryptonRichTextBox1.SelectedText);
            }

        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            var main = Form1._Form1_Instance;
            if (main == null) return;

            var activeChild = main.ActiveMdiChild;
            if (activeChild == null) return;

            // ActiveMdiChild が KryptonForm でも Form でも Controls に RichTextBox があれば取得できる
            var rtb = activeChild.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (rtb == null) return;

            //画像を貼り付け
            Clipboard.SetImage(pictureBox1.Image);
            rtb.Paste();
        }

        private void BarcodeAndQRCodeScanerDialog_Load(object sender, EventArgs e)
        {
            kryptonCheckBox1_CheckedChanged(sender, e);

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

        private void BarcodeAndQRCodeScanerDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form1._Form1_Instance.kryptonRibbonGroupButton15.Checked = false;
        }

        private void BarcodeAndQRCodeScanerDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
            GC.Collect();
        }


        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if(pictureBox1.Image != null)
                Clipboard.SetImage(pictureBox1.Image);
        }

        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.Copy();
        }

        private void kryptonContextMenuItem2_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.Cut();
        }

        private void kryptonContextMenuItem3_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.Paste();
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                kryptonContextMenu1.Show(this);
            }
        }

        private void kryptonContextMenuItem7_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.SelectAll();
        }

        private void kryptonContextMenuItem6_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {

        }

        private void kryptonButton13_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
