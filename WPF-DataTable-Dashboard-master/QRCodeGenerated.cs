using System;
using System.Drawing;
using QRCoder;

namespace DataGrid
{
    public static class QRCodeGenerated
    {
        /// <summary>
        /// 生成二维码
        /// </summary>
        /// <param name="msg">信息</param>
        /// <param name="version">版本 1 ~ 40</param>
        /// <param name="pixel">像素点大小</param>
        /// <param name="icon_path">图标路径</param>
        /// <param name="icon_size">图标尺寸</param>
        /// <param name="icon_border">图标边框厚度</param>
        /// <param name="white_edge">二维码白边</param>
        /// <returns>位图</returns>
        public static Bitmap QRCode_Generate(string msg)
        {
            int version = Convert.ToInt16(5);

            int pixel = Convert.ToInt16(100);

            int icon_size = Convert.ToInt16(20);

            int icon_border = Convert.ToInt16(1);

            string icon_path = Environment.CurrentDirectory + "\\WaterMarkPic\\QRCode_Icon.jpg";

            bool white_edge = true;

            if (msg == "")
            {
                msg = "您未输入任何文字或链接";
            }

            QRCoder.QRCodeGenerator code_generator = new QRCoder.QRCodeGenerator();

            QRCoder.QRCodeData code_data = code_generator.CreateQrCode(msg, QRCoder.QRCodeGenerator.ECCLevel.M/* 这里设置容错率的一个级别 */, true, true, QRCoder.QRCodeGenerator.EciMode.Utf8, version);

            QRCoder.QRCode code = new QRCoder.QRCode(code_data);

            Bitmap icon = new Bitmap(icon_path);

            Bitmap bmp = code.GetGraphic(pixel, Color.Black, Color.White, icon, icon_size, icon_border, white_edge);

            return bmp;

        }
    }
}
