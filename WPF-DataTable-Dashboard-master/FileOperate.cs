using iText.IO.Image;
using iText.Kernel.Pdf.Extgstate;
using iText.Kernel.Pdf;
using Spire.Doc;
using Spire.Xls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Diagnostics;

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
            Trace.WriteLine(content);
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
    }
}
