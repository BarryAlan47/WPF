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
using Spire.Doc.Fields.Shapes;
using System.Linq;

namespace DataGrid
{
    public partial class MainWindow : System.Windows.Window
    {
        ObservableCollection<Member> members = new ObservableCollection<Member>(); //待添加水印的文件列表信息，用于界面展示用
        ObservableCollection<Member> addedWaterMarkFileList = new ObservableCollection<Member>();//已添加水印的文件列表，用于界面展示用
        GetFileInfo getFileInfo = new GetFileInfo();//初始化获取文件路径的类
        FileOperate fileOperate = new FileOperate();//初始化操作文件的类
        List<string> file_list = new List<string>();//声明一个列表，用于保存待添加水印的文件列表
        TextBlock title_TextBlock;//标题文字
        TextBlock text_NoFile;//暂未选择任何文件、暂无任何添加水印记录
        System.Windows.Controls.Button button_AddFile;//
        System.Windows.Controls.Button menuButton_AddWaterMark;
        System.Windows.Controls.Button menuButton_A2a;
        System.Windows.Controls.Button menuButton_QRCodeGenerated;
        Border a2a_Panel;
        Border qRCode_Panel;
        System.Windows.Controls.TextBox textBox_a;
        System.Windows.Controls.TextBox textBox_A;
        System.Windows.Controls.Image qRCode_Image;
        System.Windows.Media.BrushConverter converter = new System.Windows.Media.BrushConverter();//改变首字符圈圈颜色用的

        public MainWindow()
        {
            InitializeComponent();
            title_TextBlock = (TextBlock)MainGrid.FindName("Title_TextBlock");
            text_NoFile = (TextBlock)FatherGrid.FindName("Text_NoFiles");
            button_AddFile = (System.Windows.Controls.Button)MainGrid.FindName("Button_AddFile");
            menuButton_AddWaterMark = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_AddWaterMark");
            menuButton_A2a = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_A2a");
            menuButton_QRCodeGenerated = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_QRCodeGenerated");
            a2a_Panel = (Border)MainGrid.FindName("A2a_Panel");
            qRCode_Panel = (Border)MainGrid.FindName("QRCode_Panel");
            textBox_a = (System.Windows.Controls.TextBox)a2a_Panel.FindName("TextBox_a");
            textBox_A = (System.Windows.Controls.TextBox)a2a_Panel.FindName("TextBox_A");
            qRCode_Image = (System.Windows.Controls.Image)QRCode_Panel.FindName("QRCode_Image");
            file_list.Clear();
        }

        private bool IsMaximize = false;
        /// <summary>
        /// 双击窗口边缘最大化界面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// 按住窗口边缘可以拖动窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }
        /// <summary>
        /// 【暂时没有选择任何文件文字】展示
        /// </summary>
        /// <param name="flag"></param>
        public void show_NoFile_Text(Boolean flag)
        {
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
        private void FillGridData() 
        {
            //获取日志行数
            int loglines = fileOperate.GetLogFileLines();

            //获取未添加水印的文件的数量
            int UnWaterMarkFileCount = 0;
            foreach (var file in members)
            {
                if (file.Flag)
                {
                    UnWaterMarkFileCount++;
                }
            }

            //打开系统窗口获取文件路径列表，相同的文件则忽略
            file_list = getFileInfo.GetFilePath();

            for (int i = 0; i < file_list.Count; i++)
            {
                bool flag = true;

                foreach (var file in members)
                {
                    if (file.FilePath == file_list[i])
                    {
                        flag = false;
                        break;
                    }
                }
                if (flag)
                {
                    Hashtable fileFullInfo = getFileInfo.GetFileFullInfo(file_list[i]);
                    members.Add(new Member
                    {
                        FilePath = fileFullInfo["filePath"].ToString(),
                        Number = (members.Count + i + 1).ToString(),
                        Character = fileFullInfo["fileName"].ToString()[..1],
                        BgColor = (System.Windows.Media.Brush)fileFullInfo["bgColor"],
                        FileName = fileFullInfo["fileName"].ToString(),
                        FileDir = fileFullInfo["fileDir"].ToString(),
                        AddWaterMarkDate = fileFullInfo["addWaterMarkDate"].ToString(),
                        FileType = fileFullInfo["fileType"].ToString(),
                        Flag = false
                    });
                }
            }
            Trace.WriteLine("--------------------file_list.Count:" + file_list.Count);
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
        /// 获取文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetFilePathButton_Click(object sender, RoutedEventArgs e)
        {
            FillGridData();
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

            int waitForAddWaterMarkFileStartIndex = 0;
            int waitForAddWaterMarkFileCount = 0;

            foreach (var file in members)
            {
                if (!file.Flag) 
                {
                    waitForAddWaterMarkFileCount++;
                }
            }
            waitForAddWaterMarkFileStartIndex = members.Count - waitForAddWaterMarkFileCount;
            //Trace.WriteLine("------------------------file_list.Count:" + file_list.Count);
            //Trace.WriteLine("------------------------waitForAddWaterMarkFileStartIndex:" + waitForAddWaterMarkFileStartIndex);
            //Trace.WriteLine("------------------------waitForAddWaterMarkFileCount:" + waitForAddWaterMarkFileCount);
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
                        FillGridData();
                        break;
                }
            }
            else
            {
                List<int> addedWaterMarkRowIndex = new List<int>();
                addedWaterMarkRowIndex.Clear();
                Task task = Task.Run(() =>
                {
                    this.Dispatcher.Invoke(new Action(() => {
                        addingWaterMark_Mask.Visibility = Visibility.Visible;
                        addingWaterMark_Icon.Visibility = Visibility.Visible;
                        addingWaterMark_TextBox.Text = "请稍等，正在添加水印中...(0" + "/" + waitForAddWaterMarkFileCount + ")";
                    }));

                    int AddedWaterMarkFileCount = 0;

                    for (int i = waitForAddWaterMarkFileStartIndex; i < members.Count; i++)
                    {
                        AddedWaterMarkFileCount++;

                        Hashtable fileFullInfo = getFileInfo.GetFileFullInfo(members[i].FilePath);

                        string filePath = fileFullInfo["filePath"].ToString();
                        string fileDir = fileFullInfo["fileDir"].ToString();
                        string fileName = fileFullInfo["fileName"].ToString();
                        string fileExtension = fileFullInfo["fileExtension"].ToString();
                        string fileType = fileFullInfo["fileType"].ToString();
                        System.Windows.Media.Brush bgColor = (System.Windows.Media.Brush)fileFullInfo["bgColor"];

                        fileOperate.StartAddWaterMark(fileOperate, filePath, fileDir, fileName, fileExtension);

                        this.Dispatcher.Invoke(new Action(() =>
                        {
                            addingWaterMark_TextBox.Text = "请稍等，正在添加水印中...(" + AddedWaterMarkFileCount + "/" + waitForAddWaterMarkFileCount + ")";
                        })
                        );
                        members[i].Flag = true;
                        Trace.WriteLine("members[" + i + "].Flag:" + members[i].Flag);
                        string logInfo = (loglines + AddedWaterMarkFileCount) + "|"+ fileName.Substring(0, 1)  + "|" + fileName + "|" + fileDir + "|" + fileFullInfo["addWaterMarkDate"].ToString() + "|"+ fileType + "|"+ filePath;
                        fileOperate.LogsWriter(logInfo);
                    }
                    this.Dispatcher.Invoke(new Action(() =>
                    {
                        addingWaterMark_TextBox.Text = "文件已全部添加水印！";
                        addingWaterMark_Icon.Kind = MahApps.Metro.IconPacks.PackIconMaterialKind.CheckBold;
                        addingWaterMark_Icon.Foreground = (System.Windows.Media.Brush)converter.ConvertFromString("#FF42D12F");
                        TimeDelay.Delay(1000);
                        addingWaterMark_Mask.Visibility = Visibility.Collapsed; 
                        addingWaterMark_Icon.Visibility = Visibility.Collapsed;
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
            string fileDir = getFileInfo.GetFileDir(filePath_Old);
            string fileName = getFileInfo.GetFileName(filePath_Old);
            string fileExtension = System.IO.Path.GetExtension(filePath_Old);
            string filePath = fileDir + "\\" + fileName + "(已添加水印)" + fileExtension;
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
            text_NoFile.Text = "暂未选择任何文件";
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
        /// 已添加水印文件夹TabButton点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabButton_AddedWaterMarkFile_Click(object sender, RoutedEventArgs e)
        {
            addedWaterMarkFileList.Clear();
            string LogContent = fileOperate.LogsReader();
            string[] list_LogInfo = fileOperate.ReadLogInfoByLine();
            System.Windows.Media.Brush bgColor;
            if (LogContent != "\n") 
            {
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
                    addedWaterMarkFileList.Add(new Member
                    {
                        Number = logInfo[0],
                        Character = logInfo[1],
                        BgColor = bgColor,
                        FileName = logInfo[2],
                        FileDir = logInfo[3],
                        AddWaterMarkDate = logInfo[4],
                        FileType = logInfo[5]
                    });
                }
            }
            AddedWatermarkFile_Grid.ItemsSource = addedWaterMarkFileList;
            membersDataGrid.Visibility = Visibility.Collapsed;
            AddedWatermarkFile_Grid.Visibility = Visibility.Visible;
            text_NoFile.Text = "暂无任何添加水印的操作记录";
            if (addedWaterMarkFileList.Count != 0)
            {
                show_NoFile_Text(false);
            }
            else 
            {
                show_NoFile_Text(true);
            }
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
            DataGridTemplateColumn templeColumn = membersDataGrid.Columns[4] as DataGridTemplateColumn;
            if (templeColumn == null)
            {
                Trace.WriteLine("---------------------templeColumn == null------------------");
            }
            else 
            {
                Trace.WriteLine("---------------------templeColumn is not null------------------");
            }
            object item = membersDataGrid.CurrentCell.Item;
            FrameworkElement element = templeColumn.GetCellContent(item);
            System.Windows.Controls.Button expander = (System.Windows.Controls.Button)templeColumn.CellTemplate.FindName("System.Windows.Controls.Button", element);
            //expander.Visibility = Visibility.Collapsed;
            this.Dispatcher.Invoke(new Action(() =>
            {
                expander.Visibility = Visibility.Collapsed;
            }));
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
        /// <summary>
        /// 将转换后的大写人民币数值拷贝到剪贴板中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Copy_A_Value_Button_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Clipboard.SetDataObject(textBox_A.Text);
        }
        /// <summary>
        /// 点击生成二维码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void QRCode_Generate_Button_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox qRCode_URL = (System.Windows.Controls.TextBox)QRCode_Panel.FindName("QRCode_URL");

            string str_msg = qRCode_URL.Text;

            Bitmap bmp = QRCodeGenerated.QRCode_Generate(str_msg);

            IntPtr hBitmap = bmp.GetHbitmap();

            ImageSource wpfBitmap = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            

            qRCode_Image.Source = wpfBitmap;
        }
        /// <summary>
        /// 点击保存二维码到本地指定位置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
    /// <summary>
    /// DataGrid成员
    /// </summary>
    public class Member
    {
        public string FilePath { get; set; }
        public string Character { get; set; }
        public System.Windows.Media.Brush BgColor { get; set; }
        public string Number { get; set; }
        public string FileName { get; set; }
        public string FileDir { get; set; }
        public string AddWaterMarkDate { get; set; }
        public string FileType { get; set; }
        public bool Flag { get; set; }
    }
}