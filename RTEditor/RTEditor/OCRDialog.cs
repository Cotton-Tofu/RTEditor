using ComponentFactory.Krypton.Toolkit;
using Microsoft.WindowsAPICodePack.Dialogs;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Advanced;
using SixLabors.ImageSharp.Formats.Bmp;
using SixLabors.ImageSharp.Processing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Tesseract;
using Windows.AI.MachineLearning;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;
using Windows.UI.Xaml.Media.Imaging;
using static System.Net.Mime.MediaTypeNames;
using Image = SixLabors.ImageSharp.Image;


namespace RTEditor
{
    public partial class OCRDialog : KryptonForm
    {
        //タイトルバーをダークモードにするためのAPI
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, int[] attrValue, int attrSize);

        public OCRDialog()
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

        private void kryptonPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        //ビットマップ画像をグレースケール・コントラスト調節用のメソッド
        public static Bitmap ToGray(Bitmap img)
        {
            var gray = new Bitmap(img.Width, img.Height);
            using (var g = Graphics.FromImage(gray))
            {
                var cm = new System.Drawing.Imaging.ColorMatrix(new float[][]
                {
            new float[]{0.299f,0.299f,0.299f,0,0},
            new float[]{0.587f,0.587f,0.587f,0,0},
            new float[]{0.114f,0.114f,0.114f,0,0},
            new float[]{0,0,0,1,0},
            new float[]{0,0,0,0,1}
                });

                var ia = new ImageAttributes();
                ia.SetColorMatrix(cm);

                g.DrawImage(img, new System.Drawing.Rectangle(0, 0, img.Width, img.Height),
                    0, 0, img.Width, img.Height, GraphicsUnit.Pixel, ia);
            }
            return gray;
        }

        public static Bitmap AdjustContrast(Bitmap image, float contrast)
        {
            // 0〜255 → 0〜1 → 中心(0.5)を基準に引き伸ばす
            float t = 0.5f * (1f - contrast);

            float[][] ptsArray =
            {
        new float[] { contrast, 0, 0, 0, 0 },
        new float[] { 0, contrast, 0, 0, 0 },
        new float[] { 0, 0, contrast, 0, 0 },
        new float[] { 0, 0, 0, 1f, 0 },
        new float[] { t, t, t, 0, 1f }
    };

            var cm = new System.Drawing.Imaging.ColorMatrix(ptsArray);
            var ia = new ImageAttributes();
            ia.SetColorMatrix(cm);

            var newBmp = new Bitmap(image.Width, image.Height);

            using (var g = Graphics.FromImage(newBmp))
            {
                g.DrawImage(
                    image,
                    new System.Drawing.Rectangle(0, 0, newBmp.Width, newBmp.Height),
                    0, 0, image.Width, image.Height,
                    GraphicsUnit.Pixel,
                    ia
                );
            }

            return newBmp;
        }




    string fileName = null;
        private async void kryptonButton1_Click(object sender, EventArgs e)
        {

            string exePath = System.Windows.Forms.Application.ExecutablePath;
            string progPath = Path.GetDirectoryName(exePath) ?? "";
            string tessdataPath = Path.Combine(progPath, "ocr_lang_jp"); // フォルダを渡す


            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "PNGファイル(*.png)|*.png|JPGファイル(*.jpg)|*.jpg|EXIF情報付きJPGファイル(*.jpg)|*.jpg|JPEGファイル(*.jpeg)|*.jpeg|GIFファイル(*.gif)|*.gif|ビットマップファイル(*.bmp)|*.bmp|TIFファイル(*.tif)|*.tif|TIFFファイル(*.tiff)|*.tiff|Windows 用アイコンファイル(*.ico)|*.ico|Windows メタファイル(*.wmf)|*.wmf|Windows 拡張メタファイル(*.emf)|*.emf|WEBPファイル(*.webp)|*.webp|TGAファイル(*.tga)|*.tga|PBMファイル(*.pbm)|*.pbm|すべてのファイル(*.*)|*.*", Title = "読み込む画像ファイルを選択...", Multiselect = false })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //高画質化処理(ピクセル数を上げビットマップ形式で保存)

                    //OCRでサポートされていない画像ファイルでもビットマップ形式に強制的に変換して保存する
                    using (SixLabors.ImageSharp.Image img = SixLabors.ImageSharp.Image.Load(openFileDialog.FileName))
                    {
                        img.SaveAsBmp(openFileDialog.FileName + ".bmp");
                        
                    }

                    //変換後のビットマップファイルを読み込み、ピクセル数二倍上昇&グレースケール&コントラスト上昇し再度保存する
                    Bitmap originalImage = new Bitmap(openFileDialog.FileName + ".bmp");
                    Bitmap reSizeImage = new Bitmap(originalImage, new System.Drawing.Size(originalImage.Width * 2, originalImage.Height * 2));
                    var gray = ToGray(reSizeImage);
                    var highContrast = AdjustContrast(gray, 1.7f);
                    reSizeImage.Save(openFileDialog.FileName+"2.bmp");

                    //下記で変換後のビットマップを使ってOCR処理をする

                    if (kryptonComboBox1.SelectedIndex == 0)
                    {
                        using (LoadingDialog loadingDialog = new LoadingDialog())
                        {
                            loadingDialog.DialogMessage = "画像スキャン中です。しばらくお待ちください...";
                            loadingDialog.Show();
                            await Task.Run(() =>
                            {
                                try
                                {
                                    using (Tesseract.TesseractEngine tesseractEngine = new Tesseract.TesseractEngine(tessdataPath, "jpn"))
                                    using (Tesseract.Pix pix = Tesseract.Pix.LoadFromFile(openFileDialog.FileName + "2.bmp"))
                                    using (Tesseract.Page pg = tesseractEngine.Process(pix))
                                    {
                                        this.Text = "光学文字認識 - " + openFileDialog.FileName;
                                        fileName = openFileDialog.FileName;
                                        Bitmap bitmap = new Bitmap(openFileDialog.FileName);
                                        pictureBox1.Image = bitmap;
                                        kryptonRichTextBox1.Text = pg.GetText();
                                        

                                        kryptonButton1.Enabled = true;
                                        kryptonButton9.Enabled = true;
                                        kryptonButton11.Enabled = true;
                                        kryptonButton4.Enabled = true;
                                        kryptonButton5.Enabled = true;
                                        kryptonButton8.Enabled = true;
                                        kryptonButton7.Enabled = true;
                                        kryptonButton2.Enabled = true;
                                        kryptonButton6.Enabled = true;
                                        kryptonButton3.Enabled = true;
                                        kryptonRichTextBox1.ReadOnly = false;


                                        if (kryptonCheckBox3.Checked == true)
                                        {
                                            kryptonRichTextBox1.SelectAll();
                                            string s = kryptonRichTextBox1.SelectedText.Replace(" ", "");
                                            kryptonRichTextBox1.SelectedText = s;
                                            kryptonRichTextBox1.DeselectAll();
                                        }


                                    }

                                }
                                catch (Exception ex)
                                {
                                    loadingDialog.Hide();
                                    TaskDialog taskDialog = new TaskDialog();
                                    taskDialog.InstructionText = "OCRの画像スキャン中にエラーが発生しました";
                                    taskDialog.Text = "内容:\n" + ex.Message + "\n\nOCRに対応していないファイルか対応言語外の画像ファイルなどを読み込もうとしている可能性があります。ほかのファイルを選択してください。またこのエラーが何度も表示される場合はサポートまで相談してください。";
                                    taskDialog.Caption = "エラー";
                                    taskDialog.OwnerWindowHandle = this.Handle;
                                    taskDialog.Icon = TaskDialogStandardIcon.Warning;
                                    taskDialog.Show();

                                    if (pictureBox1.Image == null)
                                    {
                                        kryptonButton1.Enabled = true;
                                        kryptonButton9.Enabled = false;
                                        kryptonButton11.Enabled = false;
                                        kryptonButton4.Enabled = false;
                                        kryptonButton5.Enabled = false;
                                        kryptonButton8.Enabled = false;
                                        kryptonButton7.Enabled = false;
                                        kryptonButton2.Enabled = false;
                                        kryptonButton6.Enabled = false;
                                        kryptonButton3.Enabled = false;
                                        kryptonRichTextBox1.ReadOnly = true;
                                    }

                                    taskDialog.Dispose();

                                }
                                finally
                                {

                                    loadingDialog.Close();
                                    loadingDialog.Dispose();

                                    originalImage.Dispose();
                                    reSizeImage.Dispose();

                                    //高画質化したビットマップファイルは不要なのでビットマップ版のファイルがあれば削除する
                                    if (File.Exists(openFileDialog.FileName + ".bmp"))
                                    {
                                        File.Delete(openFileDialog.FileName + ".bmp");
                                        File.Delete(openFileDialog.FileName + "2.bmp");
                                    }

                                }

                            });


                        }
                     }
                    else if (kryptonComboBox1.SelectedIndex == 1)
                    {
                        using (LoadingDialog loadingDialog = new LoadingDialog())
                        {
                            loadingDialog.DialogMessage = "画像スキャン中です。しばらくお待ちください...";
                            loadingDialog.Show();
                            try
                            {

                                byte[] imageBytes = System.IO.File.ReadAllBytes(openFileDialog.FileName + "2.bmp");

                                IBuffer buffer = imageBytes.AsBuffer();

                                SoftwareBitmap softwareBitmap;
                                using (var stream = new InMemoryRandomAccessStream())
                                {
                                    await stream.WriteAsync(buffer);
                                    stream.Seek(0);
                                    var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream);
                                    softwareBitmap = await decoder.GetSoftwareBitmapAsync();
                                }

                                var language = new Windows.Globalization.Language("ja");
                                OcrEngine ocrEngine = OcrEngine.TryCreateFromLanguage(language); //OcrEngine.TryCreateFromUserProfileLanguages();

                                OcrResult ocrResult = await ocrEngine.RecognizeAsync(softwareBitmap);


                                this.Text = "光学文字認識 - " + openFileDialog.FileName;
                                fileName = openFileDialog.FileName;
                                Bitmap bitmap = new Bitmap(openFileDialog.FileName);
                                pictureBox1.Image = bitmap;
                                kryptonRichTextBox1.Text = ocrResult.Lines.Select(line => line.Text).Aggregate((current, next) => current + Environment.NewLine + next);

                                kryptonButton1.Enabled = true;
                                kryptonButton9.Enabled = true;
                                kryptonButton11.Enabled = true;
                                kryptonButton4.Enabled = true;
                                kryptonButton5.Enabled = true;
                                kryptonButton8.Enabled = true;
                                kryptonButton7.Enabled = true;
                                kryptonButton2.Enabled = true;
                                kryptonButton6.Enabled = true;
                                kryptonButton3.Enabled = true;
                                kryptonRichTextBox1.ReadOnly = false;


                                if (kryptonCheckBox3.Checked == true)
                                {
                                    kryptonRichTextBox1.SelectAll();
                                    string s = kryptonRichTextBox1.SelectedText.Replace(" ", "");
                                    kryptonRichTextBox1.SelectedText = s;
                                    kryptonRichTextBox1.DeselectAll();
                                }


                            }
                            catch (Exception ex)
                            {
                                loadingDialog.Hide();
                                TaskDialog taskDialog = new TaskDialog();
                                taskDialog.InstructionText = "OCRの実行中にエラーが発生しました";
                                taskDialog.Text = "内容:\n" + ex.Message + "\n\nOCRに対応していないファイルか対応言語外の画像ファイルなどを読み込もうとしている可能性があります。ほかのファイルを選択してください。またこのエラーが何度も表示される場合はサポートまで相談してください。";
                                taskDialog.Caption = "エラー";
                                taskDialog.OwnerWindowHandle = this.Handle;
                                taskDialog.Icon = TaskDialogStandardIcon.Warning;
                                taskDialog.Show();

                                if (pictureBox1.Image == null)
                                {
                                    kryptonButton1.Enabled = true;
                                    kryptonButton9.Enabled = false;
                                    kryptonButton11.Enabled = false;
                                    kryptonButton4.Enabled = false;
                                    kryptonButton5.Enabled = false;
                                    kryptonButton8.Enabled = false;
                                    kryptonButton7.Enabled = false;
                                    kryptonButton2.Enabled = false;
                                    kryptonButton6.Enabled = false;
                                    kryptonButton3.Enabled = false;
                                    kryptonRichTextBox1.ReadOnly = true;
                                }

                                taskDialog.Dispose();
                            }
                            finally
                            {
                                loadingDialog.Close();
                                loadingDialog.Dispose();

                                originalImage.Dispose();
                                reSizeImage.Dispose();

                                //高画質化したビットマップファイルは不要なのでビットマップ版のファイルがあれば削除する
                                if (File.Exists(openFileDialog.FileName + ".bmp"))
                                {
                                    File.Delete(openFileDialog.FileName + ".bmp");
                                    File.Delete(openFileDialog.FileName + "2.bmp");
                                }

                            }
                        }
                    }



                }
            }

        }



        private async void kryptonButton9_Click(object sender, EventArgs e)
        {
            using (LoadingDialog loadingDialog = new LoadingDialog())
            {
                kryptonButton1.Enabled = true;
                kryptonButton9.Enabled = false;
                kryptonButton11.Enabled = false;
                kryptonButton4.Enabled = false;
                kryptonButton5.Enabled = false;
                kryptonButton8.Enabled = false;
                kryptonButton7.Enabled = false;
                kryptonButton2.Enabled = false;
                kryptonButton6.Enabled = false;
                kryptonButton3.Enabled = false;
                kryptonRichTextBox1.ReadOnly = true;
                kryptonRichTextBox1.Text = string.Empty;

                loadingDialog.DialogMessage = "画像スキャン中です。しばらくお待ちください...";
                loadingDialog.Show();
                loadingDialog.Refresh();

                {
                    //高画質化処理(ピクセル数を上げビットマップ形式で保存)

                    //OCRでサポートされていない画像ファイルでもビットマップ形式に強制的に変換して保存する
                    using (SixLabors.ImageSharp.Image img = SixLabors.ImageSharp.Image.Load(fileName))
                    {
                        img.SaveAsBmp(fileName + ".bmp");
                    }

                    //変換後のビットマップファイルを読み込み、ピクセル数二倍上昇%グレースケール&コントラスト上昇し再度保存する
                    Bitmap originalImage = new Bitmap(fileName + ".bmp");
                    Bitmap reSizeImage = new Bitmap(originalImage, new System.Drawing.Size(originalImage.Width * 2, originalImage.Height * 2));
                    var gray = ToGray(reSizeImage);
                    var highContrast = AdjustContrast(gray, 1.7f);
                    reSizeImage.Save(fileName + "2.bmp");

                    //下記で変換後のビットマップを使ってOCR処理をする


                    //TesseractOCR
                    if (kryptonComboBox1.SelectedIndex == 0)
                    {
                        try
                        {
                            await Task.Run(() =>
                            {
                                string exePath = System.Windows.Forms.Application.ExecutablePath;
                                string progPath = Path.GetDirectoryName(exePath) ?? "";
                                string tessdataPath = Path.Combine(progPath, "ocr_lang_jp"); // フォルダを渡す


                                using (Tesseract.TesseractEngine tesseractEngine = new Tesseract.TesseractEngine(tessdataPath, "jpn"))
                                using (Tesseract.Pix pix = Tesseract.Pix.LoadFromFile(fileName + "2.bmp"))
                                using (Tesseract.Page pg = tesseractEngine.Process(pix, PageSegMode.Auto))
                                {
                                    this.Text = "光学文字認識 - " + fileName;
                                    Bitmap bitmap = new Bitmap(fileName);
                                    pictureBox1.Image = bitmap;
                                    kryptonRichTextBox1.Text = pg.GetText();

                                    kryptonButton1.Enabled = true;
                                    kryptonButton9.Enabled = true;
                                    kryptonButton11.Enabled = true;
                                    kryptonButton4.Enabled = true;
                                    kryptonButton5.Enabled = true;
                                    kryptonButton8.Enabled = true;
                                    kryptonButton7.Enabled = true;
                                    kryptonButton2.Enabled = true;
                                    kryptonButton6.Enabled = true;
                                    kryptonButton3.Enabled = true;
                                    kryptonRichTextBox1.ReadOnly = false;


                                    if (kryptonCheckBox3.Checked == true)
                                    {
                                        kryptonRichTextBox1.SelectAll();
                                        string s = kryptonRichTextBox1.SelectedText.Replace(" ", "");
                                        kryptonRichTextBox1.SelectedText = s;
                                        kryptonRichTextBox1.DeselectAll();
                                    }


                                }
                            });
                        }
                        catch (Exception ex)
                        {
                            loadingDialog.Hide();
                            TaskDialog taskDialog = new TaskDialog();
                            taskDialog.InstructionText = "OCRの実行中にエラーが発生しました";
                            taskDialog.Text = "内容:\n" + ex.Message + "\n\nOCRに対応していないファイルか対応言語外の画像ファイルなどを読み込もうとしている可能性があります。ほかのファイルを選択してください。またこのエラーが何度も表示される場合はサポートまで相談してください。";
                            taskDialog.Caption = "エラー";
                            taskDialog.OwnerWindowHandle = this.Handle;
                            taskDialog.Icon = TaskDialogStandardIcon.Warning;
                            taskDialog.Show();

                            if (pictureBox1.Image == null)
                            {
                                kryptonButton1.Enabled = true;
                                kryptonButton9.Enabled = false;
                                kryptonButton11.Enabled = false;
                                kryptonButton4.Enabled = false;
                                kryptonButton5.Enabled = false;
                                kryptonButton8.Enabled = false;
                                kryptonButton7.Enabled = false;
                                kryptonButton2.Enabled = false;
                                kryptonButton6.Enabled = false;
                                kryptonButton3.Enabled = false;
                                kryptonRichTextBox1.ReadOnly = true;
                            }

                            taskDialog.Dispose();
                        }
                        finally
                        {
                            loadingDialog.Close();
                            loadingDialog.Dispose();

                            originalImage.Dispose();
                            reSizeImage.Dispose();

                            //高画質化したビットマップファイルは不要なのでビットマップ版のファイルがあれば削除する
                            if (File.Exists(fileName + ".bmp"))
                            {
                                File.Delete(fileName + ".bmp");
                                File.Delete(fileName + "2.bmp");
                            }

                        }
                    }
                    //Microsoft OCR
                    else if (kryptonComboBox1.SelectedIndex == 1)
                    {
                        try
                        {
                            byte[] imageBytes = System.IO.File.ReadAllBytes(fileName + "2.bmp");

                            IBuffer buffer = imageBytes.AsBuffer();

                            SoftwareBitmap softwareBitmap;
                            using (var stream = new InMemoryRandomAccessStream())
                            {
                                await stream.WriteAsync(buffer);
                                stream.Seek(0);
                                var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream);
                                softwareBitmap = await decoder.GetSoftwareBitmapAsync();
                            }

                            var language = new Windows.Globalization.Language("ja");
                            OcrEngine ocrEngine = OcrEngine.TryCreateFromLanguage(language); //OcrEngine.TryCreateFromUserProfileLanguages();

                            OcrResult ocrResult = await ocrEngine.RecognizeAsync(softwareBitmap);


                            kryptonRichTextBox1.Text = ocrResult.Lines.Select(line => line.Text).Aggregate((current, next) => current + Environment.NewLine + next);

                            kryptonButton1.Enabled = true;
                            kryptonButton9.Enabled = true;
                            kryptonButton11.Enabled = true;
                            kryptonButton4.Enabled = true;
                            kryptonButton5.Enabled = true;
                            kryptonButton8.Enabled = true;
                            kryptonButton7.Enabled = true;
                            kryptonButton2.Enabled = true;
                            kryptonButton6.Enabled = true;
                            kryptonButton3.Enabled = true;
                            kryptonRichTextBox1.ReadOnly = false;


                            if (kryptonCheckBox3.Checked == true)
                            {
                                kryptonRichTextBox1.SelectAll();
                                string s = kryptonRichTextBox1.SelectedText.Replace(" ", "");
                                kryptonRichTextBox1.SelectedText = s;
                                kryptonRichTextBox1.DeselectAll();
                            }

                        }
                        catch (Exception ex)
                        {
                            loadingDialog.Hide();
                            TaskDialog taskDialog = new TaskDialog();
                            taskDialog.InstructionText = "OCRの実行中にエラーが発生しました";
                            taskDialog.Text = "内容:\n" + ex.Message + "\n\nOCRに対応していないファイルか対応言語外の画像ファイルなどを読み込もうとしている可能性があります。ほかのファイルを選択してください。またこのエラーが何度も表示される場合はサポートまで相談してください。";
                            taskDialog.Caption = "エラー";
                            taskDialog.OwnerWindowHandle = this.Handle;
                            taskDialog.Icon = TaskDialogStandardIcon.Warning;
                            taskDialog.Show();

                            if (pictureBox1.Image == null)
                            {
                                kryptonButton1.Enabled = true;
                                kryptonButton9.Enabled = false;
                                kryptonButton11.Enabled = false;
                                kryptonButton4.Enabled = false;
                                kryptonButton5.Enabled = false;
                                kryptonButton8.Enabled = false;
                                kryptonButton7.Enabled = false;
                                kryptonButton2.Enabled = false;
                                kryptonButton6.Enabled = false;
                                kryptonButton3.Enabled = false;
                                kryptonRichTextBox1.ReadOnly = true;
                            }

                            taskDialog.Dispose();
                        }
                        finally
                        {
                            loadingDialog.Close();
                            loadingDialog.Dispose();

                            originalImage.Dispose();
                            reSizeImage.Dispose();

                            //高画質化したビットマップファイルは不要なのでビットマップ版のファイルがあれば削除する
                            if (File.Exists(fileName + ".bmp"))
                            {
                                File.Delete(fileName + ".bmp");
                                File.Delete(fileName + "2.bmp");
                            }


                        }


                    }



                }



            }
        }
        private void kryptonPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton11_Click(object sender, EventArgs e)
        {
            this.Text = "光学文字認識";
            fileName = null;
            pictureBox1.Image = null;
            kryptonRichTextBox1.ResetText();

            kryptonButton1.Enabled = true;
            kryptonButton9.Enabled = false;
            kryptonButton11.Enabled = false;
            kryptonButton4.Enabled = false;
            kryptonButton5.Enabled = false;
            kryptonButton8.Enabled = false;
            kryptonButton7.Enabled = false;
            kryptonButton2.Enabled = false;
            kryptonButton6.Enabled = false;
            kryptonButton3.Enabled = false;
            kryptonRichTextBox1.ReadOnly = true;
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(kryptonCheckBox1.Checked == true)
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
            if(kryptonCheckBox2.Checked == true)
            {
                this.Opacity = 0.8D;
            }
            else
            {
                this.Opacity = 100;
            }
        }

        private void OCRDialog_Load(object sender, EventArgs e)
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

        private void kryptonButton10_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void kryptonSplitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {
            
            this.Refresh();
        }

        private void kryptonCheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            if(kryptonCheckBox3.Checked == true)
            {
                kryptonRichTextBox1.SelectAll();
                string s = kryptonRichTextBox1.SelectedText.Replace(" ", "");
                kryptonRichTextBox1.SelectedText = s;
                kryptonRichTextBox1.DeselectAll();
            }
            else
            {
                if(fileName != null)
                {
                    if (File.Exists(fileName) == true)
                    {
                        kryptonButton9_Click(sender, e);
                    }
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

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "リッチテキストファイル(*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt", Title = "保存する場所を選択..."})
            {
                if(saveFileDialog.ShowDialog() == DialogResult.OK)
                    kryptonRichTextBox1.SaveFile(saveFileDialog.FileName);
            }
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
                Clipboard.SetImage(pictureBox1.Image);
        }



        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "PNGファイル(*.png)|*.png|JPGファイル(*.jpg)|*.jpg|EXIF情報付きJPGファイル(*.jpg)|*.jpg|JPEGファイル(*.jpeg)|*.jpeg|GIFファイル(*.gif)|*.gif|ビットマップファイル(*.bmp)|*.bmp|TIFファイル(*.tif)|*.tif|TIFFファイル(*.tiff)|*.tiff|Windows 用アイコンファイル(*.ico)|*.ico|Windows メタファイル(*.wmf)|*.wmf|Windows 拡張メタファイル(*.emf)|*.emf|WEBPファイル(*.webp)|*.webp|TGAファイル(*.tga)|*.tga|PBMファイル(*.pbm)|*.pbm|すべてのファイル(*.*)|*.*", Title = "画像ファイルを保存する場所を選択...", FileName = fileName })
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
                            using (SixLabors.ImageSharp.Image image = Image.Load(fileName))
                            {
                                image.SaveAsWebp(saveFileDialog.FileName);
                            }
                        }
                        //TGAファイル
                        else if (saveFileDialog.FilterIndex == 13)
                        {
                            using (SixLabors.ImageSharp.Image image = Image.Load(fileName))
                            {
                                image.SaveAsTga(saveFileDialog.FileName);
                            }
                        }
                        //PBMファイル
                        else if (saveFileDialog.FilterIndex == 14)
                        {
                            using (SixLabors.ImageSharp.Image image = Image.Load(fileName))
                            {
                                image.SaveAsPbm(saveFileDialog.FileName);
                            }
                        }
                        //PBMファイル
                        else if (saveFileDialog.FilterIndex == 14)
                        {
                            using (SixLabors.ImageSharp.Image image = Image.Load(fileName))
                            {
                                image.SaveAsPbm(saveFileDialog.FileName);
                            }
                        }
                        //その他のファイル
                        else if (saveFileDialog.FilterIndex == 15)
                        {
                            using (SixLabors.ImageSharp.Image image = Image.Load(fileName))
                            {
                                image.Save(saveFileDialog.FileName);
                            }
                        }


                    }
            }
        }

        private void OCRDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (File.Exists(fileName + ".bmp"))
            {
                File.Delete(fileName + ".bmp");
            }
            if (File.Exists(fileName + "2.bmp"))
            {
                File.Delete(fileName + "2.bmp");
            }

            Form1._Form1_Instance.kryptonRibbonGroupButton14.Checked = false;

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
            if(e.Button == MouseButtons.Right)
            {
                kryptonContextMenu1.Show(this);
            }
        }

        //Bitmapを削除するため、ウィンドウを閉じた後に一度ガベージコレクションをする
        private void OCRDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
            GC.Collect();
        }

        private void kryptonContextMenuItem7_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.SelectAll();
        }

        private void kryptonRichTextBox1_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            Process.Start(e.LinkText);
        }
    }
}