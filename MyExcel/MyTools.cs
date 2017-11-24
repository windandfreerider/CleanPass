using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;


namespace HS.ExcelExt
{
    public static class MyTools
    {
        /// <summary>
        /// 显示或隐藏Excel活动工作表的批注信息
        /// </summary>
        /// <param name="sheet">活动工作表</param>
        /// <param name="YN">是或否</param>
        public static void ShowOrHideComments(this Excel.Worksheet sheet, bool YN)
        {
            for (int i = 1; i <= sheet.Comments.Count; i++)
            {
                sheet.Comments[i].Visible = YN;
            }
        }

   
    }
}
