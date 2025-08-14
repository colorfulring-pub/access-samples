using System;
using System.Runtime.InteropServices;
using AccessPIA = Microsoft.Office.Interop.Access;

namespace AccessCOMSample
{
    /// <summary>
    /// AccessアプリケーションをCOM操作するクラス
    /// </summary>
    public class AccessApp : IDisposable
    {
        // Accessアプリケーションのインスタンス
        private AccessPIA.Application app;

        // Disposeが二重に呼ばれないようにするフラグ
        private bool disposed;

        // フォームコレクションプロパティ
        public Microsoft.Office.Interop.Access.Forms Forms => app.Forms;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public AccessApp()
        {
            app = new AccessPIA.Application();
        }

        /// <summary>
        /// ファイナライザ
        /// </summary>
        ~AccessApp() => this.Dispose(false);

        /// <summary>
        /// IDisposableの実装
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disposeの本体
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                // マネージドリソースの解放（現状なし）
            }

            // アンマネージドリソースの解放
            if (app != null)
            {
                try
                {
                    // Accessアプリケーションを保存せず終了
                    app.Quit(AccessPIA.AcQuitOption.acQuitSaveAll);
                }
                catch { }
                finally
                {
                    // COMオブジェクトの解放
                    Marshal.FinalReleaseComObject(app);
                    app = null;
                }
            }

            disposed = true;
        }

        /// <summary>
        /// データベースを開く
        /// </summary>
        /// <param name="filePath">Accessファイルパス</param>
        public void OpenDatabase(string filePath)
        {
            app.OpenCurrentDatabase(filePath);
        }

        /// <summary>
        /// フォームを開く
        /// </summary>
        /// <param name="formName">フォーム名</param>
        public void OpenForm(string formName)
        {
            app.DoCmd.OpenForm(formName, AccessPIA.AcFormView.acDesign);
        }

        /// <summary>
        /// フォームを閉じる
        /// </summary>
        /// <param name="name">フォーム名</param>
        public void CloseForm(string name)
        {
            app.DoCmd.Close(AccessPIA.AcObjectType.acForm, name, AccessPIA.AcCloseSave.acSaveNo);
        }
    }
}
