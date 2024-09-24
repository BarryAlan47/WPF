using iText.IO.Image;
using iText.Kernel.Pdf.Extgstate;
using iText.Kernel.Pdf;
using Spire.Doc;
using Spire.Xls;
using System;
using System.Collections;
using System.IO;
using Spire.Doc.Documents;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Drawing.Printing;

namespace DataGrid
{
    public class FileOperate
    {
        
        //log文件路径
        string log_FilePath = Environment.CurrentDirectory + "\\Logs\\Logs.txt";
        
        //加载PDF水印图片
        System.Drawing.Image image = System.Drawing.Image.FromFile(Environment.CurrentDirectory + "\\WaterMarkPic\\PDFWaterMark.png");
        FileStream fs = new System.IO.FileStream(Environment.CurrentDirectory + "\\WaterMarkPic\\PDFWaterMark.png", System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //加载Doc水印图片
        PictureWatermark picture = new PictureWatermark();
        FileStream fileStream = new System.IO.FileStream(Environment.CurrentDirectory + "\\WaterMarkPic\\WordWaterMark.png", System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //加载Excel水印图片
        SkiaSharp.SKBitmap bm = SkiaSharp.SKBitmap.Decode(Environment.CurrentDirectory + "\\WaterMarkPic\\ExcelWaterMark.png");
        //获取Log文件内容
        public string LogsReader()
        {
            string content = File.ReadAllText(log_FilePath);
            //Trace.WriteLine(content);
            return content;
        }
        public string[] ReadLogInfoByLine() 
        {
            string[] logInfoByLine = File.ReadAllLines(log_FilePath);
            return logInfoByLine; 
        }
        public int GetLogFileLines()
        {
            int lines = 0;  //用来统计txt行数
            FileStream fs = new FileStream(log_FilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamReader sr = new StreamReader(fs);
            while (sr.ReadLine() != null)
            {
                lines++;
            }

            fs.Close();
            sr.Close();

            return lines;

        }
        public void LogsWriter(string newLog)
        {
            string content = File.ReadAllText(log_FilePath);
            content = content + newLog + "\n";
            File.WriteAllText(log_FilePath, content);
        }
        public void DOCWaterMark(string filePath, string fileDir)
        {
            Spire.Doc.Document document = new Spire.Doc.Document();
            //从磁盘加载 Word 文档
            document.LoadFromFile(filePath);

            picture.Scaling = 150;
            picture.IsWashout = false;
            picture.SetPicture(fileStream);

            document.Watermark = picture;

            //保存文档
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            document.SaveToFile(fileDir + "\\" + fileNameWithoutExt + "(已添加水印)" + fileExtension);
        }
        public void XLSWaterMark(string filePath, string fileDir)
        {
            //加载Excel文档并获取第一个工作表
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);
            foreach (var sheet in workbook.Worksheets)
            {
                sheet.PageSetup.BackgoundImage = bm;
            }
            //保存文档
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            workbook.SaveToFile(fileDir + "\\" + fileNameWithoutExt + "(已添加水印)" + fileExtension);
        }

        public void PDFWatermark(string filePath, string fileDir)
        {
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            String DEST = (fileDir + "\\" + fileNameWithoutExt + "(已添加水印)" + fileExtension);
            String IMG = Environment.CurrentDirectory + "\\WaterMarkPic\\PDFWaterMark.png";
            String SRC = filePath;
            iText.Kernel.Pdf.PdfDocument pdfDoc = new iText.Kernel.Pdf.PdfDocument(new PdfReader(SRC), new PdfWriter(DEST));
            iText.Layout.Document doc = new iText.Layout.Document(pdfDoc);
            ImageData img = ImageDataFactory.Create(IMG);

            float w = img.GetWidth() / 4.3f;
            float h = img.GetHeight() / 4.3f;

            PdfExtGState gs1 = new PdfExtGState().SetFillOpacity(0.5f);

            // Implement transformation matrix usage in order to scale image
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                PdfPage pdfPage = pdfDoc.GetPage(i);
                iText.Kernel.Geom.Rectangle pageSize = pdfPage.GetPageSize();
                float x = (pageSize.GetLeft() + pageSize.GetRight()) / 2;
                float y = (pageSize.GetTop() + pageSize.GetBottom()) / 2;
                iText.Kernel.Pdf.Canvas.PdfCanvas over = new iText.Kernel.Pdf.Canvas.PdfCanvas(pdfPage);
                over.SaveState();
                over.SetExtGState(gs1);
                over.AddImageWithTransformationMatrix(img, w, 0, 0, h, x - (w / 2), y - (h / 2), true);
                over.RestoreState();
            }

            doc.Close();
        }
        public void StartAddWaterMark(FileOperate fileOperate, string filePath, string fileDir, string fileName, string fileExtension)
        {
            // string fileExtension = Path.GetExtension(fileName);  
            switch (fileExtension)
            {
                case ".pdf":
                    Console.WriteLine("ExtensionCheck(): .pdf");
                    fileOperate.PDFWatermark(filePath, fileDir);
                    break;
                case ".doc":
                    Console.WriteLine("ExtensionCheck(): .doc");
                    fileOperate.DOCWaterMark(filePath, fileDir);
                    break;
                case ".docx":
                    Console.WriteLine("ExtensionCheck(): .docx");
                    fileOperate.DOCWaterMark(filePath, fileDir);
                    break;
                case ".xls":
                    Console.WriteLine("ExtensionCheck(): .xls");
                    fileOperate.XLSWaterMark(filePath, fileDir);
                    break;
                case ".xlsx":
                    Console.WriteLine("ExtensionCheck(): .xlsx");
                    fileOperate.XLSWaterMark(filePath, fileDir);
                    break;
                default:
                    Console.WriteLine("default");
                    break;
            }
        }
        /**
        * 在指定目录等分pdf
        * @param fileName  要分割的文档
        * @param pageNum   分割尺寸
        * @param desDir    分割后存储路径
        * @throws IOException
        */
        public void PDFSplitterByEquipartition(string fileName, int pageNum, string desDir)
        {
            PdfReader pdfReader = new PdfReader(fileName);
            PdfDocument pdf = new PdfDocument(pdfReader);
            string name;
            PdfWriter pdfWriter = null;
            PdfDocument pdfWriterDoc = null;

            for (int i = 1; i <= pdf.GetNumberOfPages(); i += pageNum)
            {
                name = desDir + "/" + i + ".pdf";
                pdfWriter = new PdfWriter(name);
                pdfWriterDoc = new PdfDocument(pdfWriter);
                int start = i;
                int end = Math.Min((start + pageNum - 1), pdf.GetNumberOfPages());
                //从页数第一页开始，
                pdf.CopyPagesTo(start, end, pdfWriterDoc);
                pdfWriterDoc.Close();
                pdfWriter.Close();
            }

            //关闭
            pdf.Close();
            pdfReader.Close();
        }
        /**
         * 返回自定义片段大小的文件，UUID名称命名。
         * @param fileName
         * @param startPage
         * @param endPage
         * @throws IOException
         */
        public void PDFSplitterByCustomize(string fileName, Hashtable hashtable)
        {
            //源文档
            PdfReader pdfReader = new PdfReader(fileName);
            PdfDocument pdf = new PdfDocument(pdfReader);
            //目标文档名
            string desDir = "";
            //生成目标文档
            PdfWriter pdfWriter = new PdfWriter(desDir);
            PdfDocument outPdfDocument = new PdfDocument(pdfWriter);
            int startPage = 0;
            int endPage = 0;
            //从页数第一页开始，
            pdf.CopyPagesTo(startPage, endPage, outPdfDocument);
            //关闭
            outPdfDocument.Close();
            pdfWriter.Close();
            pdf.Close();
            pdfReader.Close();
        }
        public void InitMoneyRequestDOC(int Type, string InfoText,string CostText)
        {
            //上一月
            string previous_month = DateTime.Parse(DateTime.Now.ToString("Y")).AddMonths(-1).ToString("yyyy年MM月"); 
            string fileName = null;
            //大写金额
            string COSTTEXT = null;
            if (CostText != "")
            {
                float textBox_a_value = float.Parse(CostText);
                var s = textBox_a_value.ToString("#L#E#D#C#K#E#D#C#J#E#D#C#I#E#D#C#H#E#D#C#G#E#D#C#F#E#D#C#.0B0A");
                var d = Regex.Replace(s, @"((?<=-|^)[^1-9]*)|((?'z'0)[0A-E]*((?=[1-9])|(?'-z'(?=[F-L\.]|$))))|((?'b'[F-L])(?'z'0)[0A-L]*((?=[1-9])|(?'-z'(?=[\.]|$))))", "${b}${z}");
                var r = Regex.Replace(d, ".", m => "负元空零壹贰叁肆伍陆柒捌玖空空空空空空空分角拾佰仟万亿兆京垓秭穰"[m.Value[0] - '-'].ToString());
                var final_text = "人民币" + r + "整";
                COSTTEXT = final_text;
            }
            else
            {
                COSTTEXT = "金额错误";
            }
            //户名
            string bodyParagraph_4_text = null;
            //开户行
            string bodyParagraph_5_text = null;
            //银行账号
            string bodyParagraph_6_text = null;
            if (Type == 0 || Type == 1)//联拓
            {
                bodyParagraph_4_text = "户  名：广西联拓信息技术有限公司";
                bodyParagraph_5_text = "开户行：招商银行股份有限公司南宁分行";
                bodyParagraph_6_text = "帐  号：771901921910605";
            }
            else 
            {
                bodyParagraph_4_text = "户  名：广西海纳电子科技有限公司";
                bodyParagraph_5_text = "开户行：桂林银行南宁分行";
                bodyParagraph_6_text = "账  号：6602 0000 8136 1000 10";
            }
            //创建一个Document对象
            Document doc = new Document();

            //添加section
            Section section = doc.AddSection();

            //设置页边距
            section.PageSetup.Margins.Left = 90f;
            section.PageSetup.Margins.Right = 90f;
            section.PageSetup.Margins.Top = 72f;
            section.PageSetup.Margins.Bottom = 72f;

            //添加一个段落作为标题
            Paragraph titleParagraph = section.AddParagraph();
            titleParagraph.AppendText("转账请示");

            Paragraph bodyParagraph_0 = section.AddParagraph();
            bodyParagraph_0.AppendText("");
            //添加两个段落作为正文
            Paragraph bodyParagraph_1 = section.AddParagraph();
            bodyParagraph_1.AppendText("馆领导：");


            Paragraph bodyParagraph_2 = section.AddParagraph();
            if (Type == 0)
            {
                fileName = "联拓_办公耗材_转账请示" + "_" + DateTime.Parse(DateTime.Now.ToString("Y")).ToString("yyyy_MM");
                bodyParagraph_2.AppendText("我馆因办公需要，向广西联拓信息技术有限公司购买" +
                    InfoText +
                    "等办公用品及耗材配件。" +
                    previous_month +
                    "的费用共计" +
                    COSTTEXT +
                    "（¥" +
                    CostText +
                    ".00）。现所有物品已到位使用，请财务给予转账，从部门办公耗材经费支出。");
            }
            else if (Type == 1)
            {
                fileName = "联拓_维修_转账请示" + "_" + DateTime.Parse(DateTime.Now.ToString("Y")).ToString("yyyy_MM");
                bodyParagraph_2.AppendText("我馆在" +
                    previous_month +
                    "工作中，部分" +
                    InfoText +
                    "出现故障，需要维修及更换配件，费用合计" +
                    COSTTEXT +
                    "（¥" +
                    CostText +
                    ".00）(详见清单)。请财务给予转账，从部门办公设备维修维护经费支出。");
            }
            else if (Type == 2)
            {
                fileName = "海纳_办公耗材_转账请示" + "_" + DateTime.Parse(DateTime.Now.ToString("Y")).ToString("yyyy_MM");
                bodyParagraph_2.AppendText("我馆因办公需要，" +
                    previous_month +
                    "向广西海纳电子科技有限公司购买" +
                    InfoText +
                    "等办公用品及耗材配件，费用共计" +
                    COSTTEXT +
                    "（¥" +
                    CostText +
                    ".00）（详见清单）。请财务给予转账，从部门办公耗材经费支出。");
            }
            else
            {
                fileName = "海纳_维修_转账请示" + "_" + DateTime.Parse(DateTime.Now.ToString("Y")).ToString("yyyy_MM");
                bodyParagraph_2.AppendText("我馆在" +
                    previous_month +
                    "工作中，部分" +
                    InfoText +
                    "等出现故障，需要维修及更换配件，费用合计" +
                    COSTTEXT +
                    "（¥" +
                    CostText +
                    ".00）(详见清单)。请财务给予转账，从部门办公设备维修维护经费支出。");
            }

            Paragraph bodyParagraph_3 = section.AddParagraph();
            bodyParagraph_3.AppendText("妥否，请领导审批。");

            Paragraph bodyParagraph_00 = section.AddParagraph();
            bodyParagraph_00.AppendText("");

            Paragraph bodyParagraph_4 = section.AddParagraph();
            bodyParagraph_4.AppendText(bodyParagraph_4_text);

            Paragraph bodyParagraph_5 = section.AddParagraph();
            bodyParagraph_5.AppendText(bodyParagraph_5_text);

            Paragraph bodyParagraph_6 = section.AddParagraph();
            bodyParagraph_6.AppendText(bodyParagraph_6_text);

            Paragraph bodyParagraph_000 = section.AddParagraph();
            bodyParagraph_000.AppendText("");

            Paragraph bodyParagraph_7 = section.AddParagraph();
            bodyParagraph_7.AppendText("网络和信息中心");

            Paragraph bodyParagraph_8 = section.AddParagraph();
            bodyParagraph_8.AppendText("经办人：      ");

            Paragraph bodyParagraph_9 = section.AddParagraph();
            bodyParagraph_9.AppendText(DateTime.Now.ToString("yyyy年MM月dd日"));


            //为标题段落创建样式
            ParagraphStyle style1 = new ParagraphStyle(doc);
            style1.Name = "titleStyle";
            style1.CharacterFormat.Bold = false;
            style1.CharacterFormat.TextColor = Color.Black;
            style1.CharacterFormat.FontName = "方正小标宋简体";
            style1.CharacterFormat.FontSize = 22;
            doc.Styles.Add(style1);
            titleParagraph.ApplyStyle("titleStyle");

            //为正文段落创建样式
            ParagraphStyle style2 = new ParagraphStyle(doc);
            style2.Name = "paraStyle";
            style2.CharacterFormat.FontName = "仿宋_GB2312";
            style2.CharacterFormat.FontSize = 16;
            doc.Styles.Add(style2);
            bodyParagraph_1.ApplyStyle("paraStyle");
            bodyParagraph_2.ApplyStyle("paraStyle");
            bodyParagraph_3.ApplyStyle("paraStyle");
            bodyParagraph_4.ApplyStyle("paraStyle");
            bodyParagraph_5.ApplyStyle("paraStyle");
            bodyParagraph_6.ApplyStyle("paraStyle");
            bodyParagraph_7.ApplyStyle("paraStyle");
            bodyParagraph_8.ApplyStyle("paraStyle");
            bodyParagraph_9.ApplyStyle("paraStyle");

            //设置段落的水平对齐方式
            titleParagraph.Format.HorizontalAlignment = HorizontalAlignment.Center;
            bodyParagraph_1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            bodyParagraph_2.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            bodyParagraph_3.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            bodyParagraph_4.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            bodyParagraph_5.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            bodyParagraph_6.Format.HorizontalAlignment = HorizontalAlignment.Justify;
            bodyParagraph_7.Format.HorizontalAlignment = HorizontalAlignment.Right;
            bodyParagraph_8.Format.HorizontalAlignment = HorizontalAlignment.Distribute;
            bodyParagraph_9.Format.HorizontalAlignment = HorizontalAlignment.Right;

            //设置首行缩进
            bodyParagraph_1.Format.FirstLineIndent = 0;
            bodyParagraph_2.Format.FirstLineIndent = 30;
            bodyParagraph_3.Format.FirstLineIndent = 30;
            bodyParagraph_4.Format.FirstLineIndent = 0;
            bodyParagraph_5.Format.FirstLineIndent = 0;
            bodyParagraph_6.Format.FirstLineIndent = 0;
            bodyParagraph_7.Format.FirstLineIndent = 0;
            bodyParagraph_8.Format.FirstLineIndent = 0;
            bodyParagraph_9.Format.FirstLineIndent = 0;

            //设置行间距
            bodyParagraph_2.Format.LineSpacing = 17f;
            //设置后间距
            titleParagraph.Format.AfterSpacing = 10;
            bodyParagraph_0.Format.AfterSpacing = 10;
            bodyParagraph_1.Format.AfterSpacing = 10;
            bodyParagraph_2.Format.AfterSpacing = 10;
            bodyParagraph_3.Format.AfterSpacing = 10;
            bodyParagraph_4.Format.AfterSpacing = 10;
            bodyParagraph_5.Format.AfterSpacing = 10;
            bodyParagraph_6.Format.AfterSpacing = 10;
            bodyParagraph_7.Format.AfterSpacing = 10;
            bodyParagraph_8.Format.AfterSpacing = 10;
            bodyParagraph_9.Format.AfterSpacing = 10;
            //bodyParagraph_00.Format.AfterSpacing = 10;
            //bodyParagraph_000.Format.AfterSpacing = 10;
            //保存文件
            doc.SaveToFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\水印工具Output\转账请示文档\" + fileName + ".docx", Spire.Doc.FileFormat.Docx2016);

            PrintDocument printDoc = doc.PrintDocument;

            //设置PrintController属性为StandardPrintController，用于隐藏打印进程
            printDoc.PrintController = new StandardPrintController();

            //打印文档
            printDoc.Print();
        }
    }
}
