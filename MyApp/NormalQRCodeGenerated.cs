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

namespace MyApp
{
    public class NormalQRCodeGenerated
    {
        /// <summary>
        /// 生成通用二维码
        /// </summary>
        /// <param name="msg">信息</param>
        /// <returns>位图</returns>
        public static Bitmap NormalQRCode_Generate(string msg)
        {
            int version = Convert.ToInt16(5);

            int pixel = Convert.ToInt16(100);

            int icon_size = Convert.ToInt16(20);

            int icon_border = Convert.ToInt16(10);

            string icon_path = Environment.CurrentDirectory + "\\WaterMarkPic\\QRCode_Icon.jpg";

            bool white_edge = true;

            if (msg == "")
            {
                msg = "您未输入任何文字或链接";
            }

            QRCoder.QRCodeGenerator code_generator = new QRCoder.QRCodeGenerator();

            QRCoder.QRCodeData code_data = code_generator.CreateQrCode(msg, QRCoder.QRCodeGenerator.ECCLevel.M/* 这里设置容错率的一个级别 */, true, true, QRCoder.QRCodeGenerator.EciMode.Default, version);

            QRCoder.QRCode code = new QRCoder.QRCode(code_data);

            Bitmap icon = new Bitmap(icon_path);

            Bitmap bmp = code.GetGraphic(pixel, Color.Black, Color.White, icon, icon_size, icon_border, white_edge);

            return bmp;

        }
    }
}
