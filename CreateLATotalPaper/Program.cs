using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MSWord = Microsoft.Office.Interop.Word;

using System.IO;

using System.Reflection;

namespace CreateLATotalPaper
{
    class one
    {
        public static void once()
        {

        }

    }
    class Program
    {
          static void Main(string[] args)
        {

            object path;                               //文件路径变量

            string strContent;                         //文本内容变量

            MSWord.Application wordApp;                    //Word应用程序变量

            MSWord.Document wordDoc;                   //Word文档变量

            path = @AppDomain.CurrentDomain.BaseDirectory + "Out.docx";                 //路径

            wordApp = new MSWord.ApplicationClass(); //初始化

            //如果已存在，则删除

            if (File.Exists((string)path))
            {

                File.Delete((string)path);

            }

            //由于使用的是COM库，因此有许多变量需要用Missing.Value代替

            Object Nothing = Missing.Value;

            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //图片文件的路径

            string filename = @AppDomain.CurrentDomain.BaseDirectory + "1_00";

            //要向Word文档中插入图片的位置

            object count = 14;
            object WdParagraph = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;//换一行;

            Object range;

            //定义该插入的图片是否为外部链接

            Object linkToFile = false;               //默认

            //定义要插入的图片是否随Word文档一起保存

            Object saveWithDocument = true;               //默认


            //2
            const int chapter = 6;
            const int zNoCount = 5;
            const int tNoCountL = 100;
            const int tNoCountR = 172;
            int tNum = 0;
            for (int tNoNum = tNoCountL; tNoNum <= tNoCountR; ++tNoNum)
            {
                if (false == File.Exists(@AppDomain.CurrentDomain.BaseDirectory + "chapter" + chapter.ToString() + "_" + tNoNum.ToString() + "\\choice\\" + tNoNum.ToString() + "_00"))
                    continue;
                else ++tNum;
                for (int zNoNum = 0; zNoNum <= zNoCount; ++zNoNum)
                {
                    Console.WriteLine("chapter" + chapter.ToString() + "_" + tNoNum.ToString() + "\\choice\\" + tNoNum.ToString() + "_0" + zNoNum.ToString());

                    //要向Word文档中插入图片的位置
                    filename = @AppDomain.CurrentDomain.BaseDirectory + "chapter" + chapter.ToString() + "_" + tNoNum.ToString() + "\\choice\\" + tNoNum.ToString() + "_0" + zNoNum.ToString();

                    wordApp.Selection.MoveDown(ref WdParagraph, ref count, ref Nothing);//移动焦点

                    wordApp.Selection.TypeParagraph();//插入段落

                    range = wordDoc.Paragraphs.Last.Range;

                    //使用InlineShapes.AddPicture方法插入图片

                    wordDoc.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref range);

                    wordDoc.Application.ActiveDocument.InlineShapes[wordDoc.Application.ActiveDocument.InlineShapes.Count].Width *= 2.53f;
                    wordDoc.Application.ActiveDocument.InlineShapes[wordDoc.Application.ActiveDocument.InlineShapes.Count].Height *= 2.53f;

                }
            }
            //WdSaveFormat为Word 2007文档的保存格式

            object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;

            //将wordDoc文档对象的内容保存为DOCX文档

            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //关闭wordDoc文档对象

            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);

            //关闭wordApp组件对象

            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

            Console.WriteLine(path + " 创建完毕！");
            Console.WriteLine("{0} in total", tNum);
            Console.ReadKey(true);

        }

    }
}
