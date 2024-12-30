using iText.IO.Image;
using iText.Kernel.Pdf.Extgstate;
using iText.Kernel.Pdf;
using Spire.Doc;
using Spire.Xls;
using Spire.Pdf;
using System;
using System.Collections;
using System.IO;
using Spire.Doc.Documents;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Text;
using HandyControl.Tools.Extension;
using iText.Layout.Font;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace MyApp
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

        //方正小标宋简体字体文件路径
        string FZXBSJW = Environment.CurrentDirectory + "\\Fonts\\FZXBSJW.TTF";
        //仿宋GB_2312字体文件路径
        string FSGB_2312 = Environment.CurrentDirectory + "\\Fonts\\仿宋_GB2312.ttf";
        //获取Log文件内容
        public string LogsReader()
        {
            string content = File.ReadAllText(log_FilePath);
            //Trace.WriteLine(content);
            return content;
        }
        /// <summary>
        /// 读取日志
        /// </summary>
        /// <returns></returns>
        public string[] ReadLogInfoByLine()
        {
            string[] logInfoByLine = File.ReadAllLines(log_FilePath);
            return logInfoByLine;
        }
        /// <summary>
        /// 获取日志行数
        /// </summary>
        /// <returns></returns>
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
        /// <summary>
        /// 已添加水印文件记录日志
        /// </summary>
        /// <param name="newLog"></param>
        public void LogsWriter(string newLog)
        {
            string content = File.ReadAllText(log_FilePath);
            content = content + newLog + "\n";
            File.WriteAllText(log_FilePath, content);
        }
        /// <summary>
        /// 为Doc文档添加水印
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileDir"></param>
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
        /// <summary>
        /// 为XLS表格添加水印
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileDir"></param>
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
        /// <summary>
        /// 为PDF文件添加水印
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileDir"></param>
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
        /// <summary>
        /// 为MP4视频文件添加水印
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileDir"></param>
        public void MP4WaterMark(string filePath, string fileDir)
        {
            //取得ffmpeg.exe的路径,路径配置在Web.Config中,如:<add   key="ffmpeg"   value="E:\aspx1\ffmpeg.exe"   />  
            string watermarkPath = Environment.CurrentDirectory + "\\WaterMarkPic\\VideoWaterMark.png";
            //string position = "main_w-overlay_w-10:main_h-overlay_h-10";//水印位于视频右下角
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            string outputFilePath = fileDir + "\\" + fileNameWithoutExt + "(已添加水印)" + fileExtension;

            //float opacity = 0.6f;
            //string scaledWatermarkPath = Path.Combine(Path.GetTempPath(), "scaled_watermark.png");
            //ScaleWatermarkImage(watermarkPath, scaledWatermarkPath, 0.25f);
            string ffmpeg = Environment.CurrentDirectory + "\\ffmpeg\\bin\\ffmpeg.exe";
            //建立ffmpeg进程
            System.Diagnostics.ProcessStartInfo WaterMarkstartInfo = new System.Diagnostics.ProcessStartInfo(ffmpeg);
            //后台运行
            WaterMarkstartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
            //运行参数
            string config = "   -i   " + filePath + " -vf \"movie=" + watermarkPath + " [watermark]; [in][watermark] overlay=0:0 [out]\" " + outputFilePath;
            string config2 = "-i " + filePath + " -i " + watermarkPath + "  -filter_complex \" overlay=10:10 \"  -b 1024k -acodec copy " + outputFilePath;
            Trace.WriteLine("config2:" + config2);
            WaterMarkstartInfo.Arguments = config2;
            //开始加水印
            System.Diagnostics.Process.Start(WaterMarkstartInfo);
        }
        public void VideoWaterMark(string filePath,string fileDir)
        {
            //string position = "main_w-overlay_w-10:main_h-overlay_h-10";//水印位于视频右下角
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            string outputPath = fileDir + "\\" + fileNameWithoutExt + "(已添加水印)" + fileExtension;
            string watermarkPath = Environment.CurrentDirectory + "\\WaterMarkPic\\VideoWaterMark1.png";
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("视频文件未找到：" + filePath);

                if (!File.Exists(watermarkPath))
                    throw new FileNotFoundException("水印图片未找到：" + watermarkPath);

                // 确保透明度范围在 0 到 1 之间
                float opacity = 0.6f;

                // 定义 ffmpeg 和 ffprobe 路径
                string ffmpegPath = @"C:\Users\12040\source\repos\MyApp\MyApp\bin\Release\net8.0-windows\ffmpeg\bin\ffmpeg.exe";

                if (!File.Exists(ffmpegPath))
                    throw new FileNotFoundException("ffmpeg 可执行文件未找到：" + ffmpegPath);

                // 构建命令行参数
                string arguments = $"-i \"{filePath}\" -i \"{watermarkPath}\" " +
                                   $"-filter_complex \"overlay=main_w-overlay_w-10:10\" " +
                                   $"-c:v libx264 -preset fast -crf 18 -threads {Environment.ProcessorCount} -c:a copy \"{outputPath}\"";

                // 调用 ffmpeg
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = ffmpegPath,
                        Arguments = arguments,
                        UseShellExecute = false,
                        RedirectStandardError = true, // 用于读取实时输出
                        RedirectStandardOutput = false, // 不需要标准输出
                        CreateNoWindow = true // 隐藏命令行窗口
                    }
                };

                process.Start();

                // 实时读取 FFmpeg 的错误流（显示进度）
                while (!process.StandardError.EndOfStream)
                {
                    string errorLine = process.StandardError.ReadLine();
                    if (!string.IsNullOrWhiteSpace(errorLine))
                    {
                        Trace.WriteLine(errorLine); // 打印到控制台
                    }
                }

                process.WaitForExit();

                // 检查 FFmpeg 的退出代码
                if (process.ExitCode != 0)
                {
                    throw new Exception("FFmpeg 处理失败，请检查命令参数和输入文件。");
                }

                Trace.WriteLine("水印添加成功，输出文件路径：" + outputPath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("发生错误：" + ex.Message);
            }
        }
        /// <summary>
        /// 为jpg、png图片添加水印
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="watermarkPath"></param>
        /// <param name="outputPath"></param>
        public void ImageWaterMark(string filePath, string fileDir)
        {
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            string outputPath = fileDir + "\\" + fileNameWithoutExt + "(已添加水印)" + fileExtension;
            try
            {
                // 加载目标图片
                using Bitmap targetImage = new Bitmap(filePath);

                // 加载水印图片
                using Bitmap originalWatermarkImage = new Bitmap(Environment.CurrentDirectory + "\\WaterMarkPic\\PDFWaterMark.png");
                Bitmap watermarkImage = originalWatermarkImage;

                // 确保透明度值在 0 到 1 范围内
                float opacity = 0.5f;

                // 如果水印图片比目标图片大，按比例缩放
                float scale = Math.Min(
                    (float)targetImage.Width / originalWatermarkImage.Width,
                    (float)targetImage.Height / originalWatermarkImage.Height
                );
                if (scale < 1.0f)
                {
                    int newWidth = (int)(originalWatermarkImage.Width * scale);
                    int newHeight = (int)(originalWatermarkImage.Height * scale);
                    watermarkImage = new Bitmap(originalWatermarkImage, new Size(newWidth, newHeight));
                }

                // 计算水印位置（居中）
                int centerX = (targetImage.Width - watermarkImage.Width) / 2;
                int centerY = (targetImage.Height - watermarkImage.Height) / 2;

                // 创建绘图区域
                using Bitmap resultImage = new Bitmap(targetImage.Width, targetImage.Height);
                using Graphics graphics = Graphics.FromImage(resultImage);

                // 设置高质量绘图模式
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;

                // 绘制目标图片
                graphics.DrawImage(targetImage, 0, 0, targetImage.Width, targetImage.Height);

                // 设置水印透明度
                ColorMatrix colorMatrix = new ColorMatrix
                {
                    Matrix33 = opacity // 透明度
                };

                using ImageAttributes imageAttributes = new ImageAttributes();
                imageAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

                // 绘制水印图片
                Rectangle watermarkRectangle = new Rectangle(centerX, centerY, watermarkImage.Width, watermarkImage.Height);
                graphics.DrawImage(
                    watermarkImage,
                    watermarkRectangle,
                    0,
                    0,
                    watermarkImage.Width,
                    watermarkImage.Height,
                    GraphicsUnit.Pixel,
                    imageAttributes
                );

                // 保存结果图片
                resultImage.Save(outputPath, ImageFormat.Jpeg);
                Trace.WriteLine("水印添加成功，已保存到：" + outputPath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("处理图片时出错：" + ex.Message);
            }
        }
        public void StartAddWaterMark(FileOperate fileOperate, string filePath, string fileDir, string fileName, string fileExtension)
        {
            switch (fileExtension)
            {
                case ".pdf":
                    Trace.WriteLine("ExtensionCheck(): .pdf");
                    fileOperate.PDFWatermark(filePath, fileDir);
                    break;
                case ".doc":
                    Trace.WriteLine("ExtensionCheck(): .doc");
                    fileOperate.DOCWaterMark(filePath, fileDir);
                    break;
                case ".docx":
                    Trace.WriteLine("ExtensionCheck(): .docx");
                    fileOperate.DOCWaterMark(filePath, fileDir);
                    break;
                case ".xls":
                    Trace.WriteLine("ExtensionCheck(): .xls");
                    fileOperate.XLSWaterMark(filePath, fileDir);
                    break;
                case ".xlsx":
                    Trace.WriteLine("ExtensionCheck(): .xlsx");
                    fileOperate.XLSWaterMark(filePath, fileDir);
                    break;
                case ".mp4":
                    Trace.WriteLine("ExtensionCheck(): .mp4");
                    fileOperate.VideoWaterMark(filePath, fileDir);
                    break;
                case ".jpg":
                    Trace.WriteLine("ExtensionCheck(): .jpg");
                    fileOperate.ImageWaterMark(filePath, fileDir);
                    break;
                case ".png":
                    Trace.WriteLine("ExtensionCheck(): .png");
                    fileOperate.ImageWaterMark(filePath, fileDir);
                    break;
                default:
                    Trace.WriteLine("default");
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
            iText.Kernel.Pdf.PdfDocument pdf = new iText.Kernel.Pdf.PdfDocument(pdfReader);
            //目标文档名
            string desDir = "";
            //生成目标文档
            PdfWriter pdfWriter = new PdfWriter(desDir);
            iText.Kernel.Pdf.PdfDocument outPdfDocument = new iText.Kernel.Pdf.PdfDocument(pdfWriter);
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
        /// <summary>
        /// 撰写报账文档
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="InfoText"></param>
        /// <param name="CostText"></param>
        public void InitMoneyRequestDOC(int Type, string InfoText, string CostText)
        {
            //上一月
            string previous_month = DateTime.Parse(DateTime.Now.ToString("Y")).AddMonths(-1).ToString("yyyy年MM月");
            string fileName = null;
            //大写金额
            string COSTTEXT = null;
            //if (string.IsNullOrEmpty(InfoText))
            if (InfoText == "0" || InfoText == null)
            {
                InfoText = "未正确填写耗材或维修内容";
            }

            //if (string.IsNullOrEmpty(CostText))
            if (CostText == "0" || CostText == null)
            {
                COSTTEXT = "金额错误";
            }
            else
            {
                COSTTEXT = AaConvert.a2Afunc(CostText);
                if (!CostText.Contains("."))
                {
                    CostText = CostText + ".00";
                }
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
                "）。现所有物品已到位使用，请财务给予转账，从部门办公耗材经费支出。");
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
                    "）(详见清单)。请财务给予转账，从部门办公设备维修维护经费支出。");
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
                    "）（详见清单）。请财务给予转账，从部门办公耗材经费支出。");
            }
            else
            {
                fileName = "海纳_维修_转账请示" + "_" + DateTime.Parse(DateTime.Now.ToString("Y")).ToString("yyyy_MM");
                bodyParagraph_2.AppendText("我馆在" +
                    previous_month +
                    "工作中，部分" +
                    InfoText +
                    "出现故障，需要维修及更换配件，费用合计" +
                    COSTTEXT +
                    "（¥" +
                    CostText +
                    "）(详见清单)。请财务给予转账，从部门办公设备维修维护经费支出。");
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
            bodyParagraph_8.AppendText("经办人：______");

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

            //为空白行创建样式
            ParagraphStyle style3 = new ParagraphStyle(doc);
            ParagraphStyle style4 = new ParagraphStyle(doc);
            style3.Name = "spaceStyle_1";
            style3.CharacterFormat.FontSize = 22;
            doc.Styles.Add(style3);
            style4.Name = "spaceStyle_2";
            style4.CharacterFormat.FontSize = 36;
            doc.Styles.Add(style4);
            bodyParagraph_0.ApplyStyle("spaceStyle_1");
            bodyParagraph_00.ApplyStyle("spaceStyle_2");
            bodyParagraph_000.ApplyStyle("spaceStyle_2");

            //设置段落的水平对齐方式
            titleParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            bodyParagraph_1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;
            bodyParagraph_2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;
            bodyParagraph_3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;
            bodyParagraph_4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;
            bodyParagraph_5.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;
            bodyParagraph_6.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;
            bodyParagraph_7.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
            bodyParagraph_8.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
            bodyParagraph_9.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

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

            //查找指定文本
            TextSelection[] text1 = doc.FindAllString("______", false, true);
            TextSelection[] text2 = doc.FindAllString("金额错误", false, true);
            TextSelection text3 = doc.FindString("¥", false, false);
            TextSelection text4 = doc.FindString("维修内容", false, false);

            if (text1 != null)
            {
                //更改特定文本的字体颜色
                foreach (TextSelection seletion in text1)
                {
                    seletion.GetAsOneRange().CharacterFormat.TextColor = Color.White;
                }
            }

            if (text2 != null)
            {
                foreach (TextSelection seletion in text2)
                {
                    seletion.GetAsOneRange().CharacterFormat.TextColor = Color.Red;
                }
            }
            if (text3 != null)
            {
                //text3.GetAsOneRange().CharacterFormat.FontName = "宋体";
            }
            if (text4 != null)
            {
                //text4.GetAsOneRange().CharacterFormat.TextColor = Color.Red;
            }

            //保存文件
            doc.SaveToFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\水印工具Output\转账请示文档\" + fileName + ".docx", Spire.Doc.FileFormat.Docx2016);
            Word2PDF(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\水印工具Output\转账请示文档\" + fileName + ".docx", fileName);
            //StartPrintPDF(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\水印工具Output\转账请示文档\" + fileName + ".pdf");
        }
        /// <summary>
        /// 打印Doc文件
        /// </summary>
        /// <param name="filePath"></param>
        public void StartPrintDoc(string filePath) 
        {
            //初始化Document实例
            Document doc = new Document();

            //加载一个Word文档
            doc.LoadFromFile(filePath);

            //获取PrintDocument对象
            PrintDocument printDoc = doc.PrintDocument;

            //设置PrintController属性为StandardPrintController，用于隐藏打印进程
            printDoc.PrintController = new StandardPrintController();

            //打印文档
            printDoc.Print();
        }
        /// <summary>
        /// 打印PDF文件
        /// </summary>
        /// <param name="filePath"></param>
        public void StartPrintPDF(string filePath)
        {
            //加载PDF文档
            var doc = new Spire.Pdf.PdfDocument();
            doc.LoadFromFile(filePath);

            //静默打印PDF文档
            PrintDocument printDoc = doc.PrintDocument;
            printDoc.PrintController = new StandardPrintController();
            printDoc.Print();
        }
        /// <summary>
        /// Word转PDF
        /// </summary>
        /// <param name="filePath"></param>
        public void Word2PDF(string filePath,string fileName) 
        {
            //加载文档
            Document doc = new Document(false);
            doc.LoadFromFile(filePath);

            //嵌入未安装的字体.
            ToPdfParameterList ppl = new ToPdfParameterList()
            {
                PrivateFontPaths = new List<PrivateFontPath>()
                {
                    new PrivateFontPath("方正小标宋简体", FZXBSJW),
                    new PrivateFontPath("仿宋_GB2312", FSGB_2312)
                }
            };

            //保存文档.
            doc.SaveToFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\水印工具Output\转账请示文档\" + fileName + ".pdf", ppl);
        }
        
        /// <summary>
        /// 处理接收用户输入的图片裁切参数
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public string[] ProcessString(string input)
        {
            // 检查输入是否为空或仅包含空白字符
            if (string.IsNullOrWhiteSpace(input))
                throw new ArgumentException("输入不能为空或仅包含空白字符");

            // 按 | 分割字符串
            var parts = input.Split('|');

            // 根据分割结果的长度判断处理逻辑
            switch (parts.Length)
            {
                case 5: // 已经是目标格式
                    return parts;

                case 2: // 只有一个 | 号
                        // 返回前四个相同的字符串和最后一个单独的字符串
                    return new[] { parts[0], parts[0], parts[0], parts[0], parts[1] };

                case 1: // 没有 | 号
                        // 返回前四个相同的字符串和最后一个为 "0"
                    return new[] { parts[0], parts[0], parts[0], parts[0], "0" };

                default: // 不支持其他格式
                    throw new ArgumentException("输入格式无效：支持的格式有单个值、包含1个或4个 | 的字符串");
            }
        }
        /// <summary>
        /// 图片裁切
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="leftOffset"></param>
        /// <param name="rightOffset"></param>
        /// <param name="topOffset"></param>
        /// <param name="bottomOffset"></param>
        /// <param name="cornerRadius"></param>
        public void StartPictureCropping(string filePath,string fileDir,int leftOffset,int rightOffset, int topOffset, int bottomOffset, int cornerRadius) 
        {
            Trace.WriteLine(filePath);
            Trace.WriteLine(fileDir);
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            Trace.WriteLine(fileNameWithoutExt);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            Trace.WriteLine(fileExtension);
            string savePath = fileDir + "\\" + fileNameWithoutExt + "(已裁切)" + fileExtension;
            Trace.WriteLine(savePath);
            try
            {
                using (Image originalImage = Image.FromFile(filePath))
                {
                    int newWidth = originalImage.Width - leftOffset - rightOffset;
                    int newHeight = originalImage.Height - topOffset - bottomOffset;

                    if (newWidth <= 0 || newHeight <= 0)
                    {
                        Trace.WriteLine("裁剪范围无效，裁剪后的图片大小必须为正！");
                        return;
                    }

                    // 裁切图片
                    using (Bitmap croppedImage = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(croppedImage))
                        {
                            g.DrawImage(originalImage, new Rectangle(0, 0, newWidth, newHeight),
                                new Rectangle(leftOffset, topOffset, newWidth, newHeight), GraphicsUnit.Pixel);
                        }

                        // 应用圆角
                        using (Bitmap roundedImage = ApplyRoundedCorners(croppedImage, cornerRadius))
                        {
                            roundedImage.Save(savePath, ImageFormat.Png);
                            Trace.WriteLine($"图片已裁剪并保存到：{savePath}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"处理图片时发生错误：{ex.Message}");
            }
        }
        /// <summary>
        /// 裁切图片应用到的圆角子函数
        /// </summary>
        /// <param name="image"></param>
        /// <param name="cornerRadius"></param>
        /// <returns></returns>
        static Bitmap ApplyRoundedCorners(Bitmap image, int cornerRadius)
        {
            Bitmap roundedImage = new Bitmap(image.Width, image.Height);
            using (Graphics g = Graphics.FromImage(roundedImage))
            {
                // 高质量抗锯齿设置
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                g.Clear(Color.Transparent);

                // 创建高精度路径
                using (GraphicsPath path = new GraphicsPath())
                {
                    int diameter = cornerRadius * 2;

                    // 添加圆角路径
                    path.AddArc(0, 0, diameter, diameter, 180, 90);
                    path.AddArc(image.Width - diameter, 0, diameter, diameter, 270, 90);
                    path.AddArc(image.Width - diameter, image.Height - diameter, diameter, diameter, 0, 90);
                    path.AddArc(0, image.Height - diameter, diameter, diameter, 90, 90);

                    path.CloseFigure();

                    // 创建图像掩码
                    using (Bitmap mask = new Bitmap(image.Width, image.Height))
                    {
                        using (Graphics gMask = Graphics.FromImage(mask))
                        {
                            gMask.SmoothingMode = SmoothingMode.AntiAlias;
                            gMask.Clear(Color.Transparent);

                            using (Brush maskBrush = new SolidBrush(Color.White))
                            {
                                gMask.FillPath(maskBrush, path);
                            }
                        }

                        // 应用掩码
                        using (TextureBrush brush = new TextureBrush(image))
                        {
                            g.FillPath(brush, path);
                        }
                    }
                }
            }
            return roundedImage;
        }
    }

}
