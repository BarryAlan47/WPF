using System.Collections.Generic;
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.IO;
using Spire.Doc;
using System.Windows.Forms;
using Spire.Xls;
using System.Windows.Controls;
using iText.IO.Image;
using iText.Kernel.Pdf.Extgstate;
using iText.Kernel.Pdf;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Windows.Media.Imaging;
using System.Collections;
using System.Diagnostics;

namespace DataGrid
{
    public partial class MainWindow : System.Windows.Window
    {
        ObservableCollection<Member> members = new ObservableCollection<Member>();
        ObservableCollection<Member> addedWaterMarkFileList = new ObservableCollection<Member>();
        GetFileInfo getFileInfo = new GetFileInfo();
        FileOperate fileOperate = new FileOperate();
        List<string> file_list = new List<string>();
        TextBlock title_TextBlock;
        System.Windows.Controls.Button button_AddFile;
        System.Windows.Controls.Button menuButton_AddWaterMark;
        System.Windows.Controls.Button menuButton_A2a;
        System.Windows.Controls.Button menuButton_QRCodeGenerated;
        Border a2a_Panel;
        Border qRCode_Panel;
        System.Windows.Controls.TextBox textBox_a;
        System.Windows.Controls.TextBox textBox_A;
        System.Windows.Media.BrushConverter converter = new System.Windows.Media.BrushConverter();//改变首字符圈圈颜色用的

        public MainWindow()
        {
            InitializeComponent();
            title_TextBlock = (TextBlock)MainGrid.FindName("Title_TextBlock");
            button_AddFile = (System.Windows.Controls.Button)MainGrid.FindName("Button_AddFile");
            menuButton_AddWaterMark = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_AddWaterMark");
            menuButton_A2a = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_A2a");
            menuButton_QRCodeGenerated = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_QRCodeGenerated");
            a2a_Panel = (Border)MainGrid.FindName("A2a_Panel");
            qRCode_Panel = (Border)MainGrid.FindName("QRCode_Panel");
            textBox_a = (System.Windows.Controls.TextBox)a2a_Panel.FindName("TextBox_a");
            textBox_A = (System.Windows.Controls.TextBox)a2a_Panel.FindName("TextBox_A");
        }

        private bool IsMaximize = false;
        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                if (IsMaximize)
                {
                    this.WindowState = WindowState.Normal;
                    this.Width = 1080;
                    this.Height = 720;

                    IsMaximize = false;
                }
                else
                {
                    this.WindowState = WindowState.Maximized;

                    IsMaximize = true;
                }
            }
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }
        /// <summary>
        /// 暂时没有选择任何文件文字展示
        /// </summary>
        /// <param name="flag"></param>
        public void show_NoFile_Text(Boolean flag)
        {
            TextBlock text_NoFile = (TextBlock)FatherGrid.FindName("Text_NoFiles");
            if (flag)
            {
                text_NoFile.Visibility = Visibility.Visible;
            }
            else
            {
                text_NoFile.Visibility = Visibility.Collapsed;
            }

        }
        /// <summary>
        /// 菜单栏：添加水印按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_AddWaterMark_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                button_AddFile.Visibility = Visibility.Visible;
                title_TextBlock.Text = "添加水印";
                a2a_Panel.Visibility = Visibility.Collapsed;
                qRCode_Panel.Visibility = Visibility.Collapsed;
            }));
        }
        /// <summary>
        /// 菜单栏：人民币大小写转换
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_A2a_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                button_AddFile.Visibility = Visibility.Collapsed;
                title_TextBlock.Text = "大小写转换";
                a2a_Panel.Visibility = Visibility.Visible;
                qRCode_Panel.Visibility = Visibility.Collapsed;
            }));
        }
        /// <summary>
        /// 菜单栏：生成二维码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_QRCodeGenerated_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                button_AddFile.Visibility = Visibility.Collapsed;
                title_TextBlock.Text = "二维码生成器";
                a2a_Panel.Visibility = Visibility.Collapsed;
                qRCode_Panel.Visibility = Visibility.Visible;
            }));
        }
        /// <summary>
        /// 获取文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetFilePathButton_Click(object sender, RoutedEventArgs e)
        {
            int loglines = fileOperate.GetLogFileLines();
            //获取文件前的fileList的数量
            int BeforeAddFile = file_list.Count;
            //获取选择的文件列表
            file_list = getFileInfo.GetFilePath();
            //改变首字符圈圈颜色用的
            var converter = new System.Windows.Media.BrushConverter();

            members.Clear();
            membersDataGrid.ItemsSource = null;

            for (int i = 0; i < file_list.Count; i++)
            {
                string fileName = getFileInfo.GetFileName(file_list[i]);
                string fileDir = getFileInfo.GetFileDir(file_list[i]);
                string fileType = "";
                System.Windows.Media.Brush bgColor;
                if (System.IO.Path.GetExtension(file_list[i]) == ".pdf")
                {
                    fileType = "PDF文件";
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#FF5252");
                }
                else if (System.IO.Path.GetExtension(file_list[i]) == ".doc" || System.IO.Path.GetExtension(file_list[i]) == ".docx")
                {
                    fileType = "Word文档";
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#1E88E5");
                }
                else if (System.IO.Path.GetExtension(file_list[i]) == ".xls" || System.IO.Path.GetExtension(file_list[i]) == ".xlsx")
                {
                    fileType = "Excel表格";
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#0CA678");
                }
                else
                {
                    fileType = "未知类型文件";
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#D3D3D3");
                }

                members.Add(new Member { Number = (loglines + i + 1).ToString(), Character = fileName.Substring(0, 1), BgColor = bgColor, Name = fileName, Position = fileDir, Email = "", Phone = fileType });
            }

            membersDataGrid.ItemsSource = members;

            if (members.Count != 0)
            {
                show_NoFile_Text(false);
            }
            else
            {
                show_NoFile_Text(true);
            }
        }
        /// <summary>
        /// 开始添加水印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddWaterMarkButton_Click(object sender, RoutedEventArgs e)
        {
            Border addingWaterMark_Mask = (Border)MainGrid.FindName("AddingWaterMark_Mask");
            System.Windows.Controls.TextBox addingWaterMark_TextBox = (System.Windows.Controls.TextBox)AddingWaterMark_Mask.FindName("AddingWaterMark_TextBox");
            MahApps.Metro.IconPacks.PackIconMaterial addingWaterMark_Icon = (MahApps.Metro.IconPacks.PackIconMaterial)MainGrid.FindName("AddingWaterMark_Icon");
            int loglines = fileOperate.GetLogFileLines();
            membersDataGrid.ItemsSource = members;
            if (members.Count != 0)
            {
                show_NoFile_Text(false);
            }
            else
            {
                show_NoFile_Text(true);
            }
            if (file_list.Count == 0)
            {
                var result = System.Windows.MessageBox.Show("尚未选择任何文件,您是否希望前往选择需要添加水印的文件?", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                switch (result)
                {
                    case MessageBoxResult.Cancel:
                        // User pressed Cancel
                        break;
                    case MessageBoxResult.OK:
                        // User pressed Yes
                        int BeforeAddFile = file_list.Count;

                        file_list = getFileInfo.GetFilePath();

                        members.Clear();
                        membersDataGrid.ItemsSource = null;

                        for (int i = 0; i < file_list.Count; i++)
                        {
                            string fileName = getFileInfo.GetFileName(file_list[i]);
                            string fileDir = getFileInfo.GetFileDir(file_list[i]);
                            string fileType = "";
                            System.Windows.Media.Brush bgColor;
                            if (System.IO.Path.GetExtension(file_list[i]) == ".pdf")
                            {
                                fileType = "PDF文件";
                                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#FF5252");
                            }
                            else if (System.IO.Path.GetExtension(file_list[i]) == ".doc" || System.IO.Path.GetExtension(file_list[i]) == ".docx")
                            {
                                fileType = "Word文档";
                                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#1E88E5");
                            }
                            else if (System.IO.Path.GetExtension(file_list[i]) == ".xls" || System.IO.Path.GetExtension(file_list[i]) == ".xlsx")
                            {
                                fileType = "Excel表格";
                                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#0CA678");
                            }
                            else
                            {
                                fileType = "未知类型文件";
                                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#D3D3D3");
                            }
                            members.Add(new Member { Number = (loglines + i + 1).ToString(), Character = fileName.Substring(0, 1), BgColor = bgColor, Name = fileName, Position = fileDir, Email = "", Phone = fileType });
                        }
                        membersDataGrid.ItemsSource = members;
                        break;
                }
            }
            else
            {
                Task task = Task.Run(() =>
                {
                    this.Dispatcher.Invoke(new Action(() => {
                        addingWaterMark_Mask.Visibility = Visibility.Visible;
                        addingWaterMark_Icon.Visibility = Visibility.Visible;
                        addingWaterMark_TextBox.Text = "请稍等，正在添加水印中...(0" + "/" + file_list.Count + ")";
                    }));
                    for (int i = 0; i < file_list.Count; i++)
                    {
                        string filePath = file_list[i];
                        string fileDir = getFileInfo.GetFileDir(filePath);
                        string fileName = getFileInfo.GetFileName(filePath);
                        string fileExtension = System.IO.Path.GetExtension(file_list[i]);
                        string fileType = "";
                        System.Windows.Media.Brush bgColor;
                        if (System.IO.Path.GetExtension(file_list[i]) == ".pdf")
                        {
                            fileType = "PDF文件";
                            bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#FF5252");
                        }
                        else if (System.IO.Path.GetExtension(file_list[i]) == ".doc" || System.IO.Path.GetExtension(file_list[i]) == ".docx")
                        {
                            fileType = "Word文档";
                            bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#1E88E5");
                        }
                        else if (System.IO.Path.GetExtension(file_list[i]) == ".xls" || System.IO.Path.GetExtension(file_list[i]) == ".xlsx")
                        {
                            fileType = "Excel表格";
                            bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#0CA678");
                        }
                        else
                        {
                            fileType = "未知类型文件";
                            bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#D3D3D3");
                        }

                        fileOperate.StartAddWaterMark(fileOperate, filePath, fileDir, fileName, fileExtension);

                        this.Dispatcher.Invoke(new Action(() =>
                        {
                            addingWaterMark_TextBox.Text = "请稍等，正在添加水印中...(" + (i + 1) + "/" + file_list.Count + ")";
                        })
                        );
                        addedWaterMarkFileList.Add(new Member { Number = (loglines + i + 1).ToString(), Character = fileName.Substring(0, 1), BgColor = members[i].BgColor, Name = fileName, Position = fileDir, Email = System.DateTime.Now.ToString("d"), Phone = fileType });
                        string logInfo = (loglines + i + 1) + "|"+ fileName.Substring(0, 1)  + "|" + fileName + "|" + fileDir + "|" + System.DateTime.Now.ToString("d") + "|"+ fileType + "|"+ filePath;
                        fileOperate.LogsWriter(logInfo);
                    }
                    this.Dispatcher.Invoke(new Action(() =>
                    {
                        addingWaterMark_TextBox.Text = "文件已全部添加水印！";
                        addingWaterMark_Icon.Kind = MahApps.Metro.IconPacks.PackIconMaterialKind.CheckBold;
                        addingWaterMark_Icon.Foreground = (System.Windows.Media.Brush)converter.ConvertFromString("#FF42D12F");
                        TimeDelay.Delay(1000);
                        //AddedWatermarkFile_Grid.ItemsSource = addedWaterMarkFileList;
                        addingWaterMark_Mask.Visibility = Visibility.Collapsed; 
                        addingWaterMark_Icon.Visibility = Visibility.Collapsed;
                        members.Clear();
                        membersDataGrid.ItemsSource = members;
                    })
                    ); 
                }
                );
            }
        }
        /// <summary>
        /// 移除已经添加的文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedRowIndex = membersDataGrid.SelectedIndex;
            for (int i = selectedRowIndex; i < members.Count; i++)
            {
                members[i].Number = (int.Parse(members[i].Number) - 1).ToString();
            }
            file_list.RemoveAt(selectedRowIndex);
            members.RemoveAt(selectedRowIndex);
            membersDataGrid.ItemsSource = null;
            membersDataGrid.ItemsSource = members;
            if (members.Count != 0)
            {
                show_NoFile_Text(false);
            }
            else
            {
                show_NoFile_Text(true);
            }
        }
        /// <summary>
        /// 打开已添加水印的文件路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PositionFileButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath_Old = fileOperate.ReadLogInfoByLine()[AddedWatermarkFile_Grid.SelectedIndex].Split('|')[6];
            Trace.WriteLine("-------------filePath_Old-----------:" + filePath_Old);
            string fileDir = getFileInfo.GetFileDir(filePath_Old);
            string fileName = getFileInfo.GetFileName(filePath_Old);
            string fileExtension = System.IO.Path.GetExtension(filePath_Old);
            string filePath = fileDir + "\\" + fileName + "(已添加水印)" + fileExtension;
            Trace.WriteLine("-------------filePath-----------:" + filePath);
            if (!System.IO.File.Exists(filePath)) 
            {
                Trace.WriteLine("所选的文件已被移动至其他地方");
            }
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            //string file = @"c:/ windows/notepad.exe"; 
            psi.Arguments = " /select," + filePath;
            System.Diagnostics.Process.Start(psi); 
        }
        /// <summary>
        /// 已选择文件夹TabButton点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabButton_SeletedFile_Click(object sender, RoutedEventArgs e)
        {
            membersDataGrid.Visibility = Visibility.Visible;
            AddedWatermarkFile_Grid.Visibility = Visibility.Collapsed;
        }
        /// <summary>
        /// 已添加水印文件夹TabButton点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabButton_AddedWaterMarkFile_Click(object sender, RoutedEventArgs e)
        {
            addedWaterMarkFileList.Clear();
            string[] list_LogInfo = fileOperate.ReadLogInfoByLine();
            System.Windows.Media.Brush bgColor;
            foreach (var item in list_LogInfo)
            {
                string[] logInfo = item.Split('|');
                if (logInfo[5] == "PDF文件")
                {
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#FF5252");
                }
                else if (logInfo[5] == "Word文档")
                {
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#1E88E5");
                }
                else if (logInfo[5] == "Excel表格")
                {
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#0CA678");
                }
                else
                {
                    bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#D3D3D3");
                }
                addedWaterMarkFileList.Add(new Member { Number = logInfo[0], Character = logInfo[1], BgColor = bgColor, Name = logInfo[2], Position = logInfo[3], Email = logInfo[4], Phone = logInfo[5] });
            }
            AddedWatermarkFile_Grid.ItemsSource = addedWaterMarkFileList;
            membersDataGrid.Visibility = Visibility.Collapsed;
            AddedWatermarkFile_Grid.Visibility = Visibility.Visible;
        }
        private void PageUpButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("已经是第一页");
        }
        private void PageDownButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("已经是最后一页");
        }
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            //fileOperate.LogsOperate();
        }


        private void CloseTheAppButton_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void TextBox_a_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (textBox_a.Text != "")
            {
                float textBox_a_value = float.Parse(textBox_a.Text);
                var s = textBox_a_value.ToString("#L#E#D#C#K#E#D#C#J#E#D#C#I#E#D#C#H#E#D#C#G#E#D#C#F#E#D#C#.0B0A");
                var d = Regex.Replace(s, @"((?<=-|^)[^1-9]*)|((?'z'0)[0A-E]*((?=[1-9])|(?'-z'(?=[F-L\.]|$))))|((?'b'[F-L])(?'z'0)[0A-L]*((?=[1-9])|(?'-z'(?=[\.]|$))))", "${b}${z}");
                var r = Regex.Replace(d, ".", m => "负元空零壹贰叁肆伍陆柒捌玖空空空空空空空分角拾佰仟万亿兆京垓秭穰"[m.Value[0] - '-'].ToString());
                var final_text = "人民币" + r + "整";
                textBox_A.Text = final_text;
            }
            else
            {
                textBox_A.Text = "";
            }
        }

        private void Copy_A_Value_Button_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Clipboard.SetDataObject(textBox_A.Text);
        }

        private void QRCode_Generate_Button_Click(object sender, RoutedEventArgs e)
        {
            int version = Convert.ToInt16(5);

            int pixel = Convert.ToInt16(100);

            System.Windows.Controls.TextBox qRCode_URL = (System.Windows.Controls.TextBox)QRCode_Panel.FindName("QRCode_URL");

            string str_msg = qRCode_URL.Text;

            int int_icon_size = Convert.ToInt16(20);

            int int_icon_border = Convert.ToInt16(1);

            bool b_we = true;
            if (qRCode_URL.Text == "")
            {
                str_msg = "您未输入任何文字或链接";
            }
            Bitmap bmp = QRCodeGenerated.QRCode_Generate(str_msg, version, pixel, Environment.CurrentDirectory + "\\WaterMarkPic\\QRCode_Icon.jpg", int_icon_size, int_icon_border, b_we);

            IntPtr hBitmap = bmp.GetHbitmap();
            ImageSource wpfBitmap = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            System.Windows.Controls.Image qRCode_Image = (System.Windows.Controls.Image)QRCode_Panel.FindName("QRCode_Image");

            qRCode_Image.Source = wpfBitmap;
        }

        private void SaveQRCode_Button_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Image qRCode_Image = (System.Windows.Controls.Image)QRCode_Panel.FindName("QRCode_Image");
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Image Files (*.bmp, *.png, *.jpg)|*.bmp;*.png;*.jpg | All Files | *.*";
            sfd.RestoreDirectory = true;//保存对话框是否记忆上次打开的目录 
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create((BitmapSource)qRCode_Image.Source));
                using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create))
                    encoder.Save(stream);
            }
        }

    }

    public class Member
    {
        public string Character { get; set; }
        public System.Windows.Media.Brush BgColor { get; set; }
        public string Number { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
    }
}