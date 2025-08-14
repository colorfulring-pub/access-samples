using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AccessPIA = Microsoft.Office.Interop.Access;

namespace AccessCOMSample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // デフォルトのAccessファイル名を設定
            FileNameTextBox.Text = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                , "TestDatabase1.accdb");
        }

        /// <summary>
        /// プロパティ取得
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExecuteButton_Click(object sender, EventArgs e)
        {
            string accessFileName = FileNameTextBox.Text;
            string targetFormName = "フォーム1";    // 取得対象のフォーム名

            // 結果のクリア
            ResultListBox.Items.Clear();

            try
            {
                Cursor = Cursors.WaitCursor;

                // Accessアプリケーションを開く
                using (var app = new AccessApp())
                {
                    // Accessファイルを開く
                    app.OpenDatabase(accessFileName);

                    AccessPIA.Form form = null;

                    try
                    {
                        // フォームを開いてプロパティを取得
                        app.OpenForm(targetFormName);

                        form = app.Forms[targetFormName];
                        var properties = AccessReflector.GetProperties(form);

                        // 結果リストに追加
                        ResultListBox.Items.AddRange(
                            properties
                            .OrderBy(p => p.Name)
                            .Select(p => $"{p.Name} = {p.Value?.ToString()}").ToArray());

                        // フォームを閉じる
                        app.CloseForm(targetFormName);
                    }
                    finally
                    {
                        // フォームオブジェクトを解放
                        if (form != null)
                        {
                            Marshal.ReleaseComObject(form);
                        }
                    }
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            MessageBox.Show("完了しました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
