using System;
using System.Windows.Forms;
using System.IO;

// <summary>
// C:\Users\{UserName}\AppData\Local\にある Edge WebView2 ランタイムと RTEditor の設定情報のキャッシュファイル(フォルダ)を削除するコンソールアプリ
//アンインストール時に行う
// </summary>
namespace DeleteCashFiles
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //WindowsのUIに準拠するためWindows Formsのビジュアルスタイルを有効にする
            System.Windows.Forms.Application.EnableVisualStyles();

            //削除確認のメッセージボックスを表示
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            DialogResult result = MessageBox.Show(path+"にある Edge WebView2 ランタイムと RTEditor の設定情報のキャッシュファイル(フォルダ)を削除しますか?\r\n\r\n「はい」を押すとキャッシュファイルの削除と製品のアンインストール処理が実行され、「いいえ」を押すとキャッシュファイルは削除されず製品のアンインストール処理のみ実行されます。\r\n\\r\n(このプログラムはアンインストール処理中に起動することを目的としておりを単体で起動している場合、アンインストール処理は実行されません。)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    // キャッシュファイルの削除処理
                    string edgeWebView2CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)+@"\RTEditor");
                    string rtEditorCachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)+@"\Cotton_Tofu");

                    if (System.IO.Directory.Exists(edgeWebView2CachePath))
                    {
                        Console.WriteLine("Edge WebView 2のランタイムファイルが見つかりました。");
                        System.IO.Directory.Delete(edgeWebView2CachePath, true);
                        Console.WriteLine("Edge WebView 2のランタイムファイルを削除しました。");
                    }
                    else
                    {
                        Console.WriteLine("Edge WebView 2のランタイムファイルが存在しません。");
                    }

                    if (System.IO.Directory.Exists(rtEditorCachePath))
                    {
                        Console.WriteLine("RTEditorの設定情報ファイルが見つかりました。");
                        System.IO.Directory.Delete(rtEditorCachePath, true);
                        Console.WriteLine("RTEditorの設定情報ファイルを削除しました。");
                    }
                    else
                    {
                        Console.WriteLine("RTEditorの設定情報ファイルが存在しません。");
                    }
                    MessageBox.Show("キャッシュファイルが削除されました。", "完了", MessageBoxButtons.OK);


                    //コンソールアプリを閉じる
                    Environment.Exit(0);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"キャッシュファイルの削除中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK);

                    //コンソールアプリを閉じる
                    Environment.Exit(0);
                }
            }
            else
            {
                MessageBox.Show("キャッシュファイルは削除されませんでした。", "完了", MessageBoxButtons.OK);
                //コンソールアプリを閉じる
                Environment.Exit(0);
            }

        }
    }
}

