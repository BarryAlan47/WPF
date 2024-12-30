using System;
using System.Drawing;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Net;
using System.Text;
using QRCoder;
using System.Net.Http.Headers;
using System.Net.Http;
using Org.BouncyCastle.Asn1.Crmf;
using static QRCoder.PayloadGenerator.ShadowSocksConfig;
using System.IO;
using Org.BouncyCastle.Tsp;
using System.Text.Json.Nodes;
using System.Drawing.Imaging;
using ClosedXML.Excel;
using System.Diagnostics;

namespace MyApp
{
    public class NormalQRCodeGenerated
    {
        /// <summary>
        /// 生成通用二维码
        /// </summary>
        /// <param name="pagePath"></param>
        /// <param name="saveName"></param>
        public static async Task NormalQRCode_Generate(string pagePath,string savePath)
        {
            int version = Convert.ToInt16(5);

            int pixel = Convert.ToInt16(100);

            int icon_size = Convert.ToInt16(20);

            int icon_border = Convert.ToInt16(10);

            string icon_path = Environment.CurrentDirectory + "\\WaterMarkPic\\QRCode_Icon.jpg";

            bool white_edge = true;

            if (pagePath == "")
            {
                pagePath = "您未输入任何文字或链接";
            }

            QRCoder.QRCodeGenerator code_generator = new QRCoder.QRCodeGenerator();

            QRCoder.QRCodeData code_data = code_generator.CreateQrCode(pagePath, QRCoder.QRCodeGenerator.ECCLevel.M/* 这里设置容错率的一个级别 */, true, true, QRCoder.QRCodeGenerator.EciMode.Default, version);

            QRCoder.QRCode code = new QRCoder.QRCode(code_data);

            Bitmap icon = new Bitmap(icon_path);

            Bitmap bmp = code.GetGraphic(pixel, Color.Black, Color.White, icon, icon_size, icon_border, white_edge);

            //return bmp;
            //保存到本地目录内
            try
            {
                // 保存为 PNG 格式
                bmp.Save(savePath, ImageFormat.Png);
                Console.WriteLine($"图片已保存到: {savePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存失败: {ex.Message}");
            }
            finally
            {
                // 释放资源
                bmp.Dispose();
            }

        }
        /// <summary>
        /// 批量生成通用二维码
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static async Task WXQRCodeMultiplyGenerated(string filePath)
        {
            Trace.WriteLine($"WXQRCode Multiplied: {filePath}");
            try
            {
                string excelFilePath = filePath; //  Excel 文件路径

                string outputDirectory = @"C:\Users\12040\Desktop\二维码";  // 输出目录路径

                if (!Directory.Exists(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed();

                    foreach (var row in rows.Skip(1)) // 跳过标题行
                    {
                        string SaveName = row.Cell(1).GetValue<string>();
                        string pagePath = row.Cell(2).GetValue<string>();
                        if (string.IsNullOrWhiteSpace(SaveName) || string.IsNullOrWhiteSpace(pagePath))
                            continue;

                        string outputPath = Path.Combine(outputDirectory, $"{SaveName}.png");

                        // 调用生成二维码方法
                        await NormalQRCode_Generate(pagePath, outputPath);
                    }
                }

                Trace.WriteLine("二维码批量生成完成！");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"发生错误：{ex.Message}");
            }
        }
    }
}
