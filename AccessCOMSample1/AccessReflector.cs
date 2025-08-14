using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using AccessPIA = Microsoft.Office.Interop.Access;

namespace AccessCOMSample
{
    /// <summary>
    /// プロパティ取得クラス
    /// </summary>
    public class AccessReflector
    {
        /// <summary>
        /// プロパティを取得
        /// </summary>
        /// <param name="form">対象フォームオブジェクト</param>
        /// <returns>プロパティの名称と値のタプルリスト</returns>
        public static List<(string Name, object Value)> GetProperties(AccessPIA.Form form)
        {
            var result = new List<(string, object)>();
            var type = form.GetType();

            // 公開されているインスタンスプロパティを列挙
            // 特殊な名前のプロパティや、読み取り不可のプロパティは除外
            foreach (var prop in
                type
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => !p.IsSpecialName && p.CanRead
                ))
            {
                object value = null;

                // プロパティ値を取得するGetメソッドを取得
                var getMethod = prop.GetGetMethod();
                var getMethodParameterLength = getMethod?.GetParameters().Length;

                // パラメータ付きのGetメソッドはスキップ(サンプルのため除外しています)
                if (getMethodParameterLength > 0)
                    continue;

                try
                {
                    // プロパティ値を取得
                    value = getMethod.Invoke(form, new object[] { });
                    var valueType = value?.GetType();

                    // COM オブジェクトの場合は特別な処理（サンプルのため何もしていません）
                    if (value != null && (valueType.IsImport || valueType.FullName == "System.__ComObject"))
                    {
                    }
                    else
                    {
                        // イベントプロシージャの場合は「イベント プロシージャ」として表示
                        if (value is string str && str.Equals("[Event Procedure]"))
                        {
                            value = "[イベント プロシージャ]";
                        }
                        // バイト配列の場合は「バイナリ」として表示
                        else if (value is byte[])
                        {
                            value = "(バイナリ)";
                        }
                    }

                    // 結果リストに追加
                    result.Add((prop.Name, value));
                }
                catch (Exception ex)
                {
                    // サンプルのため例外発生時何もしない
                }
                finally
                {
                    // プロパティ値がCOMオブジェクトの場合は必要に応じてリリースする
                    if (value != null && Marshal.IsComObject(value))
                    {
                        Marshal.ReleaseComObject(value);
                    }
                }
            }

            return result;
        }
    }
}
