using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Color = System.Drawing.Color;

namespace CodeDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            //获取文件名
            //遍历文件

            var codeFolder = @"D:\code\temp\NoloHomeNavigation\Nolo_WPFAssistant\";

            var directoryInfo = new System.IO.DirectoryInfo(codeFolder);
            FileInfo[] afi = directoryInfo.GetFiles("*.*", SearchOption.AllDirectories);
            string fileName, fullpath;
            IList<string> fileList = new List<string>();

            for (int i = 0; i < afi.Length; i++)
            {
                fullpath = afi[i].FullName;
                fileName = fullpath.ToLower();
                if ((fileName.EndsWith(".xaml") || fileName.EndsWith(".cs")) && !fileName.Contains("obj"))
                {
                    fileList.Add(fullpath);
                    //Console.WriteLine(fullpath.Replace(codeFolder, ""));
                }
            }

            var versinn = "1.8.17";

            using (DocX document = DocX.Create($"D:\\nolohome\\codedoc\\nolohome源码_{versinn}.docx"))
            {
                // Add a title.
                document.InsertParagraph("Nolo_Home源代码").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

                //  var headingTypes = Enum.GetValues(typeof(HeadingType));

                foreach (var filename in fileList)
                {
                    var text = filename.Replace(codeFolder, "");
                    // Add a paragraph.
                    var p = document.InsertParagraph().Append(text);
                    // Set the paragraph's heading type.
                    p.Heading(HeadingType.Heading2);


                    var pFileContent = document.InsertParagraph();

                    // Append some text and add formatting.
                    pFileContent.Append(File.ReadAllText(filename))
                        .Color(Color.Black);
                }





                document.Save();
                Console.WriteLine("\tCreated: Heading.docx\n");
            }



            Console.ReadKey();
        }
    }
}
