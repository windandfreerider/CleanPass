using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using HS.ExcelExt;
using Hs.Tools;

namespace CleanPass
{
    public partial class Ribbon1
    {
        private bool CommentState = true ;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            sheet.CleanPassword();
        }


        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton button = sender as RibbonToggleButton;
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            sheet.ShowOrHideComments(button.Checked);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            sheet.ShowOrHideComments(CommentState);
            CommentState = !CommentState ;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //string url1 = @"http://www.matrix67.com/blog/feed";
            //Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            //rng.Value = MyGrab.GetContent(url1);
            //rng.AutoFitRange(100);

            string url2 = @"http://blog.sina.com.cn/rss/1748013412.xml";
            Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            rng.Value = MyGrab.GetContent(url2);
            rng.AutoFitRange(100);
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //测试
            Excel.Application app = Globals.ThisAddIn.Application;

            string str = "";
            str = System.Convert.ToString(System.Windows.Forms.Clipboard.GetText());

            int n = 0;
            object[,] arr = new object[100, 11];
            var with_1 = new System.Text.RegularExpressions.Regex("\\|.*?\\|");
            System.Text.RegularExpressions.MatchCollection col1 = with_1.Matches(str);
            foreach (System.Text.RegularExpressions.Match mm1 in col1)
            {
                arr[n, 0] = mm1.Value;
                n = n + 1;
            }
            app.Range["a1"].get_Resize(n, 1).Value2 = arr;
        }
    }
}


