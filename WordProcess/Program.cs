using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Text.RegularExpressions;

namespace WordProcess
{
    class Program
    {
        static void Main(string[] args)
        {

            object file = Path.GetFullPath("..") + "\\test\\test.docx";//获得bin下面的test路径
            var doc = ReadDocx(file, out var word_app);
            Console.WriteLine(string.Format("总计{0}个段落", doc.Paragraphs.Count));//输出文章的段落数目
            int IndexPar = 0;
            foreach (Paragraph par in word_app.ActiveDocument.Paragraphs)
            {
                if (IsCite(par.Range.Text, out int CiteIndex))
                {
                    using (System.IO.StreamWriter fileX = new System.IO.StreamWriter(Path.GetFullPath("..") + "\\test\\demo3.txt", true))
                    {
                        if (CiteIndex != 2)
                        {
                            try
                            {
                                fileX.WriteLine(par.Range.Text.Substring(CiteIndex - 4, 8));
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }
                IndexPar++;
                Console.WriteLine(IndexPar);
            }
            Console.ReadKey();
        }


        /// <summary>
        /// 读取word文档
        /// </summary>
        /// <param name="file"></param>
        /// <param name="app"></param>
        /// <returns></returns>
        public static Document ReadDocx(object file,out Microsoft.Office.Interop.Word.Application app)
        {
            KillWnWordProcess();
            try
            {
                 app = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
                object unknow = Type.Missing;
                app.Visible = false;

                doc = app.Documents.Open(ref file,
                ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow);
                if (doc != null)
                {
                    Console.WriteLine("读取成功");
                }
                return doc;
            }
            catch (Exception ex)
            {
                app = null;
                Console.WriteLine(ex.Message);
                return null;
            }
        }
        /// <summary>
        /// 关闭Word进程
        /// </summary>
        public static void KillWnWordProcess()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }
        /// <summary>
        /// 判断是否是索引
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool IsCite(string paragraph,out int index)
        {
            Regex regex = new Regex("(?<=【)(\\d+)(?=】)");
            var result= regex.Match(paragraph);
            if (result.Success)
            {
                index = result.Index;
                return true;
            }
            else
            {
                index = 0;
                return false;
            }


        }

    }
}
