using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace MyApp
{
    class WXQRCodeGenerated
    {
        private const string AppId = "wx7f96177263663d31"; // 区博预约小程序APPID
        private const string AppSecret = "edc197317016af9411d2922a00296f2d"; // 区博预约小程序AppSecret

        private const string AppId_Main = "wxc735f8163fdb2855"; // 区博门户小程序APPID
        private const string AppSecret_Main = "1b660b6532b201595d3e85357a371c5d"; // 区博门户小程序AppSecret

        private static string _cachedAccessToken;
        private static string _cachedAccessToken_Main;
        private static DateTime _accessTokenExpiryTime;
        private static DateTime _accessTokenExpiryTime_Main;

        /// <summary>
        /// 单个生成二维码
        /// </summary>
        /// <returns></returns>
        public static async Task WXQRCodeSingleGenerated(string ChanelName, string PagePath, string SaveName)
        {
            string accessToken = null;
            if (string.IsNullOrEmpty(ChanelName))
            {
                // 获取 Access Token
                accessToken = await GetAccessTokenWithCacheAsync(AppId, AppSecret);
                if (string.IsNullOrEmpty(PagePath))
                {
                    PagePath = "pages/activity/activityDetail.html?data={\"id\":\"2a8742b650254760af8282c7d115acfe\"}";
                }
            }
            else
            {
                // 获取 Access Token
                accessToken = await GetAccessTokenWithCacheAsync(AppId_Main, AppSecret_Main);
                if (string.IsNullOrEmpty(PagePath))
                {
                    PagePath = "pages/guide/exhDetail.html?hall=" + ChanelName + "&id=fc9527a58fb172cf92d3aab1194b9790";
                }
            }

            if (string.IsNullOrEmpty(SaveName))
            {
                SaveName = "未设置保存名称";
            }
            try
            {
                string outputDirectory = @"C:\Users\12040\Desktop\水印工具Output\二维码";  // 替换为输出目录路径

                if (!Directory.Exists(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                string finalPagePath = TransformString(PagePath);
                string outputPath = Path.Combine(outputDirectory, $"{SaveName}.png");

                // 调用生成二维码方法
                await GenerateQRCodeAsync(accessToken, finalPagePath, outputPath);

                Trace.WriteLine("单个二维码生成完成！");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"发生错误：{ex.Message}");
            }
        }
        /// <summary>
        /// 批量生成二维码
        /// </summary>
        /// <returns></returns>
        public static async Task WXQRCodeMultiplyGenerated(string filePath)
        {
            Trace.WriteLine($"WXQRCode Multiplied: {filePath}");
            Trace.WriteLine("WXQRCode Multiplied2: C:\\Users\\12040\\Desktop\\测试小程序码生成表格.xlsx");
            try
            {
                string excelFilePath = filePath; //  Excel 文件路径

                string outputDirectory = @"C:\Users\12040\Desktop\水印工具Output\二维码";  // 输出目录路径

                if (!Directory.Exists(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                // 获取 Access Token
                string accessToken = await GetAccessTokenWithCacheAsync(AppId, AppSecret);
                string accessToken_Main = await GetAccessTokenWithCacheAsync(AppId_Main, AppSecret_Main);

                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed();

                    foreach (var row in rows.Skip(1)) // 跳过标题行
                    {
                        string fileName = row.Cell(1).GetValue<string>();
                        string ChanelName = row.Cell(2).GetValue<string>();
                        string pagePath = row.Cell(3).GetValue<string>();
                        string finalPagePath = null;
                        //if (string.IsNullOrWhiteSpace(finalPagePath))
                        //    continue;
                        if (string.IsNullOrWhiteSpace(fileName))
                            fileName = "Temp";


                        string outputPath = Path.Combine(outputDirectory, $"{fileName}.png");
                        if (string.IsNullOrWhiteSpace(ChanelName))
                        {
                            finalPagePath = TransformString(pagePath);
                            // 调用生成二维码方法
                            await GenerateQRCodeAsync(accessToken, finalPagePath, outputPath);
                        }
                        else
                        {
                            Trace.WriteLine("_________ChanelName__________:" + ChanelName);
                            Trace.WriteLine("_________id__________:" + pagePath);
                            finalPagePath = TransformString_2(ChanelName, pagePath);
                            // 调用生成二维码方法
                            await GenerateQRCodeAsync(accessToken_Main, finalPagePath, outputPath);

                        }
                        Trace.WriteLine("_________finalPagePath__________:" + finalPagePath);
                        

                        
                    }
                }

                Trace.WriteLine("二维码批量生成完成！");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"发生错误：{ex.Message}");
            }
        }
        /// <summary>
        /// 获取AccessToken，如果还在7200s内,则不重复获取
        /// </summary>
        /// <param name="appId"></param>
        /// <param name="appSecret"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static async Task<string> GetAccessTokenWithCacheAsync(string appId, string appSecret)
        {
            if (appId == "wx7f96177263663d31")
            {
                // 如果有缓存且未过期，直接返回缓存的区博预约小程序 access_token
                if (!string.IsNullOrEmpty(_cachedAccessToken) && DateTime.UtcNow < _accessTokenExpiryTime)
                {
                    return _cachedAccessToken;
                }
                // 否则重新获取 Access Token
                string url = $"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={appId}&secret={appSecret}";

                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    string result = await response.Content.ReadAsStringAsync();
                    var json = JsonConvert.DeserializeObject<dynamic>(result);

                    if (json?.access_token != null)
                    {
                        _cachedAccessToken = json.access_token.ToString();
                        int expiresIn = json.expires_in ?? 7200; // 默认 7200 秒有效期
                        _accessTokenExpiryTime = DateTime.UtcNow.AddSeconds(expiresIn - 60); // 提前 1 分钟过期，避免临界问题
                        return _cachedAccessToken;
                    }

                    throw new Exception($"获取 Access Token 失败：{result}");
                }
            }
            else
            {
                // 如果有缓存且未过期，直接返回缓存的区博门户小程序 access_token
                if (!string.IsNullOrEmpty(_cachedAccessToken_Main) && DateTime.UtcNow < _accessTokenExpiryTime_Main)
                {
                    return _cachedAccessToken_Main;
                }
                // 否则重新获取 Access Token
                string url = $"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={appId}&secret={appSecret}";

                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    string result = await response.Content.ReadAsStringAsync();
                    var json = JsonConvert.DeserializeObject<dynamic>(result);

                    if (json?.access_token != null)
                    {
                        _cachedAccessToken_Main = json.access_token.ToString();
                        int expiresIn = json.expires_in ?? 7200; // 默认 7200 秒有效期
                        _accessTokenExpiryTime_Main = DateTime.UtcNow.AddSeconds(expiresIn - 60); // 提前 1 分钟过期，避免临界问题
                        return _cachedAccessToken_Main;
                    }

                    throw new Exception($"获取 Access Token 失败：{result}");
                }
            }
        }
        /// <summary>
        /// 微信官方API,调用以返回二维码Buffer数据流
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="pagePath"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static async Task GenerateQRCodeAsync(string accessToken, string pagePath, string outputPath)
        {
            string url = $"https://api.weixin.qq.com/wxa/getwxacode?access_token={accessToken}";

            var payload = new
            {
                path = pagePath,
                width = 430, // 二维码宽度
                auto_color = false,
                line_color = new { r = 0, g = 0, b = 0 },
                is_hyaline = false // 设置背景是否透明
            };

            using (HttpClient client = new HttpClient())
            {
                StringContent content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(url, content);

                if (!response.IsSuccessStatusCode)
                {
                    string error = await response.Content.ReadAsStringAsync();
                    throw new Exception($"生成二维码失败：{error}");
                }

                // 保存二维码图像
                byte[] qrCodeData = await response.Content.ReadAsByteArrayAsync();
                File.WriteAllBytes(outputPath, qrCodeData);
            }
        }
        /// <summary>
        /// 【活动预约】截取字符串，生成正确的PagePath格式
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        static string TransformString(string input)
        {
            if (input.StartsWith("pages"))
            {
                // 移除 ".html"
                return input.Replace(".html", "");
            }
            else if (input.StartsWith("http"))
            {
                // 查找 id= 后面的部分
                int idIndex = input.IndexOf("id=") + 3;
                if (idIndex > 2) // 确保找到了 "id="
                {
                    string idValue = input.Substring(idIndex);
                    return $"pages/activity/activityDetail?data={{\"id\":\"{idValue}\"}}";
                }
            }

            // 不符合要求的输入返回原字符串
            return input;
        }

        /// <summary>
        /// 【AR导览】截取字符串，生成正确的PagePath格式
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        static string TransformString_2(string Channel,string id)
        {
            return $"pages/guide/exhDetail?hall=" + Channel + "&id=" + id;
        }
    }
}