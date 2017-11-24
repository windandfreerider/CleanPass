using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace HS.ExcelExt
{
    public static class MyRange
    {
        public static void AutoFitRange(this Excel.Range rng,int width=88)
        {
            rng.ColumnWidth = width;
            rng.Columns.AutoFit();
        }
     }
}
