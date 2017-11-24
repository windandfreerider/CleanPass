using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Hs.Tools
{
    public static class MyGrab
    {
        public static string GetContent(string url)
        {
            XElement xml = XElement.Load(url);

            string txt = "\r\n";
            var list = xml.Element("channel").Elements("item")
                .Select((m, index1) => txt += index1.ToString() + ":" + m.Element("title").Value + "\r\n")
                .Where((n, index2) => index2 < 10)
                .ToList();

            //返回结果文本
            return txt;
        }

        public static string GetConts(string url)
        {
            XElement xml = XElement.Load(url);
            string txt = "\r\n";
            var list = xml.Element("channel").Elements("item")
                .Select((m, index1) => txt += index1.ToString() + ":" + m.Element("title").Value + "\r\n")
                .Where((n, index2) => index2 < 5)
                .ToList();
            return txt;
        }
    }
}
