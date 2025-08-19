using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Collections;
using System.Diagnostics;
using HandyControl.Controls;
using System.Windows.Media.Animation;
using System.Text;
using System.ComponentModel;
using System.IO;

namespace MyApp
{
    public partial class MainWindow : System.Windows.Window
    {
        GetFileInfo getFileInfo = new GetFileInfo();//初始化获取文件路径的类
        FileOperate fileOperate = new FileOperate();//初始化操作文件的类

        System.Windows.Media.Brush tabButton_BorderBrush_Color_Default;
        System.Windows.Media.Brush tabButton_BorderBrush_Color_Seleted;
        System.Windows.Media.Brush tabButton_Foreground_Color_Default;
        System.Windows.Media.Brush tabButton_Foreground_Color_Seleted;
        System.Windows.Media.BrushConverter converter = new System.Windows.Media.BrushConverter();//改变首字符圈圈颜色用的

        //杂项
        MahApps.Metro.IconPacks.PackIconMaterial openOrCLoseMenuButton_Icon;
        CircleProgressBar myCircleProgressBar;
        Border p1_Image_Border;//功能头像
        ImageBrush featuredImage;

        //添加水印相关组件及列表配置
        ObservableCollection<Member> members = new ObservableCollection<Member>(); //待添加水印的文件列表信息，用于界面展示用
        ObservableCollection<Member> addedWaterMarkFileList = new ObservableCollection<Member>();//已添加水印的文件列表，用于界面展示用
        List<string> file_list = new List<string>();//声明一个列表，用于保存待添加水印的文件列表
        TextBlock title_TextBlock;//标题文字
        TextBlock text_NoFile;//暂未选择任何文件、暂无任何添加水印记录
        System.Windows.Controls.Button button_AddFile;//添加文件按钮
        System.Windows.Controls.Button addWaterMarkButton;//开始添加水印按钮
        System.Windows.Controls.Button tabButton_SelectedFiles;//标签页按钮：已选择的文件列表
        System.Windows.Controls.Button tabButton_AddedWaterMarkFiles;//标签页按钮：已添加过水印的文件列表
        MahApps.Metro.IconPacks.PackIconMaterial addingWaterMark_Icon;
        DataGridTemplateColumn templeColumn;//动态生成的列表

        //图片裁切的相关组件及列表配置
        System.Windows.Controls.Button button_SelectedFiles_PictureCropping;
        System.Windows.Controls.Button button_Start_PictureCropping;
        System.Windows.Controls.TextBox cuttingParameters_TextBox;
        ObservableCollection<Member> members_PictureCropping = new ObservableCollection<Member>();//待裁切的图片列表信息，用于界面展示用
        DataGridTemplateColumn templeColumn_PictureCropping;//动态生成的列表

        //撰写转账文档的相关组件及配置
        System.Windows.Controls.TextBox lT_Haocai_Info_TextBox;
        System.Windows.Controls.TextBox lT_Haocai_Cost_TextBox;
        System.Windows.Controls.TextBox lT_Weixiu_Info_TextBox;
        System.Windows.Controls.TextBox lT_Weixiu_Cost_TextBox;
        System.Windows.Controls.TextBox hN_Haocai_Info_TextBox;
        System.Windows.Controls.TextBox hN_Haocai_Cost_TextBox;
        System.Windows.Controls.TextBox hN_Weixiu_Info_TextBox;
        System.Windows.Controls.TextBox hN_Weixiu_Cost_TextBox;

        //通用二维码的相关组件及列表配置
        System.Windows.Controls.Button button_SelectedFiles_NormalQRCode;
        System.Windows.Controls.Button button_Start_Generate_NormalQRCode;
        System.Windows.Controls.TextBox textBox_PagePath_NormalQRCode;
        System.Windows.Controls.TextBox textBox_SaveName_NormalQRCode;
        System.Windows.Controls.Grid normalQRCode_PagePath_TextBox_Grid;
        System.Windows.Controls.Grid normalQRCode_SaveName_TextBox_Grid;
        ObservableCollection<Member> members_NormalQRCode = new ObservableCollection<Member>();//待生成太阳码的Excel列表数组信息，用于界面展示用
        DataGridTemplateColumn templeColumn_NormalQRCode;//动态生成的列表对象

        //微信太阳码的相关组件及列表配置
        System.Windows.Controls.Button button_SelectedFiles_WXQRCode;
        System.Windows.Controls.Button button_Single_WXQRCode;
        System.Windows.Controls.Button button_Multiply_WXQRCode;
        System.Windows.Controls.TextBox textBox_ChanelName;
        System.Windows.Controls.TextBox textBox_PagePath;
        System.Windows.Controls.TextBox textBox_SaveName;
        System.Windows.Controls.Grid wXQRCode_ChanelName_TextBox_Grid;
        System.Windows.Controls.Grid wXQRCode_PagePath_TextBox_Grid;
        ObservableCollection<Member> members_WXQRCode = new ObservableCollection<Member>();//待生成太阳码的Excel列表数组信息，用于界面展示用
        DataGridTemplateColumn templeColumn_WXQRCode;//动态生成的列表对象

        //人民币大小写转换的相关组件及配置列表
        System.Windows.Controls.Button menuButton_A2a;//侧边菜单栏功能按钮
        System.Windows.Controls.TextBox a2A_textBox_a;//转换前金额文本框
        System.Windows.Controls.TextBox a2A_textBox_A;//转换后金额文本框
        MahApps.Metro.IconPacks.PackIconMaterial a2A_Icon;//展示动效用的icon
        System.Windows.Media.Brush a2AIcon_bgColor_Default;//展示动效用的颜色1
        System.Windows.Media.Brush a2AIcon_bgColor_Truning;//展示动效用的颜色2


        
        public MainWindow()
        {
            InitializeComponent();
            title_TextBlock = (TextBlock)MainGrid.FindName("Title_TextBlock");
            text_NoFile = (TextBlock)FatherGrid.FindName("Text_NoFiles");
            button_AddFile = (System.Windows.Controls.Button)MainGrid.FindName("Button_AddFile");
            addWaterMarkButton = (System.Windows.Controls.Button)MainGrid.FindName("AddWaterMarkButton");
            tabButton_SelectedFiles = (System.Windows.Controls.Button)MenuButton_Grid.FindName("TabButton_SelectedFiles");
            tabButton_AddedWaterMarkFiles = (System.Windows.Controls.Button)MenuButton_Grid.FindName("TabButton_AddedWaterMarkFiles");
            file_list.Clear();
            templeColumn = membersDataGrid.Columns[4] as DataGridTemplateColumn;
            
            tabButton_BorderBrush_Color_Default = (System.Windows.Media.Brush)converter.ConvertFromString("#DAE2EA");
            tabButton_BorderBrush_Color_Seleted = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");
            tabButton_Foreground_Color_Default = (System.Windows.Media.Brush)converter.ConvertFromString("#FF121518");
            tabButton_Foreground_Color_Seleted = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");
            myCircleProgressBar = (CircleProgressBar)MainGrid.FindName("MyCircleProgressBar");
            addingWaterMark_Icon = (MahApps.Metro.IconPacks.PackIconMaterial)MainGrid.FindName("AddingWaterMark_Icon");
            openOrCLoseMenuButton_Icon = (MahApps.Metro.IconPacks.PackIconMaterial)FatherGrid.FindName("OpenOrCLoseMenuButton_Icon");
            p1_Image_Border = (Border)MenuButton_Grid.FindName("P1_Image_Border");
            featuredImage = (ImageBrush)MenuButton_Grid.FindName("FeaturedImage");

            //图片裁切相关组件实例化
            button_SelectedFiles_PictureCropping = (System.Windows.Controls.Button)PictureCroppingGrid.FindName("PictureCropping_Button_AddFile");
            button_Start_PictureCropping = (System.Windows.Controls.Button)PictureCroppingGrid.FindName("PictureCropping_Button_Start");
            cuttingParameters_TextBox = (System.Windows.Controls.TextBox)CuttingParameters_TextBox_Grid.FindName("CuttingParameters_TextBox");
            templeColumn_PictureCropping = PictureCropping_membersDataGrid.Columns[4] as DataGridTemplateColumn;

            //撰写转账请示文档相关组件实例化
            //System.Windows.Controls.TextBox lT_Haocai_Info_TextBox = (System.Windows.Controls.TextBox)LT_Haocai_Info_Grid.FindName("LT_Haocai_Info_TextBox");
            //System.Windows.Controls.TextBox lT_Haocai_Cost_TextBox = (System.Windows.Controls.TextBox)LT_Haocai_Cost_Grid.FindName("LT_Haocai_Cost_TextBox");
            //System.Windows.Controls.TextBox lT_Weixiu_Info_TextBox = (System.Windows.Controls.TextBox)LT_Weixiu_Info_Grid.FindName("LT_Weixiu_Info_TextBox");
            //System.Windows.Controls.TextBox lT_Weixiu_Cost_TextBox = (System.Windows.Controls.TextBox)LT_Weixiu_Cost_Grid.FindName("LT_Weixiu_Cost_TextBox");
            //System.Windows.Controls.TextBox hN_Haocai_Info_TextBox = (System.Windows.Controls.TextBox)HN_Haocai_Info_Grid.FindName("HN_Haocai_Info_TextBox");
            //System.Windows.Controls.TextBox hN_Haocai_Cost_TextBox = (System.Windows.Controls.TextBox)HN_Haocai_Cost_Grid.FindName("HN_Haocai_Cost_TextBox");
            //System.Windows.Controls.TextBox hN_Weixiu_Info_TextBox = (System.Windows.Controls.TextBox)HN_Weixiu_Info_Grid.FindName("HN_Weixiu_Info_TextBox");
            //System.Windows.Controls.TextBox hN_Weixiu_Cost_TextBox = (System.Windows.Controls.TextBox)HN_Weixiu_Cost_Grid.FindName("HN_Weixiu_Cost_TextBox");

            //通用二维码相关组件实例化
            button_SelectedFiles_NormalQRCode = (System.Windows.Controls.Button)NormalQRCodeGrid.FindName("NormalQRCode_Button_AddFile");
            button_Start_Generate_NormalQRCode = (System.Windows.Controls.Button)NormalQRCodeGrid.FindName("NormalQRCode_Start_Generate_Button");
            textBox_PagePath_NormalQRCode = (System.Windows.Controls.TextBox)NormalQRCode_PagePath_TextBox_Grid.FindName("NormalQRCode_PagePath_TextBox");
            textBox_SaveName_NormalQRCode = (System.Windows.Controls.TextBox)NormalQRCode_SaveName_TextBox_Grid.FindName("NormalQRCode_SaveName_TextBox");
            normalQRCode_PagePath_TextBox_Grid = (System.Windows.Controls.Grid)WXQRCodeGrid.FindName("NormalQRCode_PagePath_TextBox_Grid");
            normalQRCode_SaveName_TextBox_Grid = (System.Windows.Controls.Grid)WXQRCodeGrid.FindName("NormalQRCode_SaveName_TextBox_Grid");
            templeColumn_NormalQRCode = NormalQRCode_membersDataGrid.Columns[4] as DataGridTemplateColumn;

            //微信太阳码相关组件实例化
            button_SelectedFiles_WXQRCode = (System.Windows.Controls.Button)WXQRCodeGrid.FindName("WXQRCode_Button_AddFile");
            button_Single_WXQRCode = (System.Windows.Controls.Button)WXQRCodeGrid.FindName("WXQRCode_Button_Single");
            button_Multiply_WXQRCode = (System.Windows.Controls.Button)WXQRCodeGrid.FindName("WXQRCode_Button_Multiply");
            textBox_ChanelName = (System.Windows.Controls.TextBox)WXQRCode_ChanelName_TextBox_Grid.FindName("WXQRCode_ChanelName_TextBox");
            textBox_PagePath = (System.Windows.Controls.TextBox)WXQRCode_PagePath_TextBox_Grid.FindName("WXQRCode_PagePath_TextBox");
            textBox_SaveName = (System.Windows.Controls.TextBox)WXQRCode_SaveName_TextBox_Grid.FindName("WXQRCode_SaveName_TextBox");
            wXQRCode_ChanelName_TextBox_Grid = (System.Windows.Controls.Grid)WXQRCodeGrid.FindName("WXQRCode_ChanelName_TextBox_Grid");
            wXQRCode_PagePath_TextBox_Grid = (System.Windows.Controls.Grid)WXQRCodeGrid.FindName("WXQRCode_PagePath_TextBox_Grid");
            templeColumn_WXQRCode = WXQRCode_membersDataGrid.Columns[4] as DataGridTemplateColumn;

            //人民币大小写转换相关组件实例化
            menuButton_A2a = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_A2a");
            a2A_textBox_a = (System.Windows.Controls.TextBox)a2A_TextBox_a_Grid.FindName("_a2A_TextBox_a");
            a2A_textBox_A = (System.Windows.Controls.TextBox)a2A_TextBox_A_Grid.FindName("_a2A_TextBox_A");
            a2A_Icon = (MahApps.Metro.IconPacks.PackIconMaterial)a2A_Grid.FindName("_a2A_Icon");
            a2AIcon_bgColor_Default = (System.Windows.Media.Brush)converter.ConvertFromString("#FFA5A5A5");
            a2AIcon_bgColor_Truning = (System.Windows.Media.Brush)converter.ConvertFromString("#FF6EA1F3");
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
        /// 填充已选择文件列表(添加水印)
        /// </summary>
        private void FillGridData()
        {
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
            int CurrentFileCount = members.Count;
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
                        Number = (CurrentFileCount + i + 1).ToString(),
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
            //Trace.WriteLine("--------------------file_list.Count:" + file_list.Count);
            //Trace.WriteLine("--------------------members.Count:" + members.Count);
            //membersDataGrid.ItemsSource = null;
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
        /// 获取文件(添加水印)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_AddFile_Click(object sender, RoutedEventArgs e)
        {
            FillGridData();
            tabButton_SelectedFiles.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Button.ClickEvent));
        }
        /// <summary>
        /// 点击此按钮开始添加水印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddWaterMarkButton_Click(object sender, RoutedEventArgs e)
        {
            Border addingWaterMark_Mask = (Border)MainGrid.FindName("AddingWaterMark_Mask");
            System.Windows.Controls.TextBox addingWaterMark_TextBox = (System.Windows.Controls.TextBox)AddingWaterMark_Mask.FindName("AddingWaterMark_TextBox");
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
            if (waitForAddWaterMarkFileCount != 0)
            {
                Task task = Task.Run(() =>
                {
                    this.Dispatcher.Invoke(new Action(() =>
                    {
                        addingWaterMark_Mask.Visibility = Visibility.Visible;
                        addingWaterMark_Icon.Visibility = Visibility.Collapsed;
                        myCircleProgressBar.Value = 0;
                        myCircleProgressBar.Text = "0";
                        addingWaterMark_TextBox.Text = "请稍等，正在添加水印中...(0" + "/" + waitForAddWaterMarkFileCount + ")";
                        myCircleProgressBar.Maximum = waitForAddWaterMarkFileCount;
                    }));
                    int AddedWaterMarkFileCount = 0;
                    float ProgressBar_CurrentValue = 0f;
                    ShowAddingWaterMarkMask(true);
                    TimeDelay.Delay(800);
                    for (int i = waitForAddWaterMarkFileStartIndex; i < members.Count; i++)
                    {
                        AddedWaterMarkFileCount++;
                        ShowProgressBar(i);

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
                        ProgressBar_CurrentValue += 1;
                        this.Dispatcher.Invoke(new Action(() =>
                        {
                            myCircleProgressBar.Value = ProgressBar_CurrentValue;
                            myCircleProgressBar.Text = (i + 1).ToString();
                        })
                        );
                        members[i].Flag = true;
                        string logInfo = (loglines + AddedWaterMarkFileCount) + "|" + fileName.Substring(0, 1) + "|" + fileName + "|" + fileDir + "|" + fileFullInfo["addWaterMarkDate"].ToString() + "|" + fileType + "|" + filePath;
                        fileOperate.LogsWriter(logInfo);
                        ShowOpenFileButton(i);
                    }
                    this.Dispatcher.Invoke(new Action(() =>
                    {
                        addingWaterMark_TextBox.Text = "文件已全部添加水印！";
                        addingWaterMark_Icon.Visibility = Visibility.Visible;
                        myCircleProgressBar.Visibility = Visibility.Collapsed;
                        TimeDelay.Delay(1000);
                        ShowAddingWaterMarkMask(false);
                        TimeDelay.Delay(800);
                        //addingWaterMark_Mask.Visibility = Visibility.Collapsed;
                        addingWaterMark_Icon.Visibility = Visibility.Collapsed;
                        myCircleProgressBar.Visibility = Visibility.Visible;
                    })
                    );
                }
                );
            }
            else
            {
                if (members.Count == 0)
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
                    var result = System.Windows.MessageBox.Show("当前列表的文件均已添加水印,您是否希望前往选择需要新的需要添加水印的文件?", "提示", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
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
            }
        }
        /// <summary>
        /// 移除已经添加的文件(添加水印)
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
                //Trace.WriteLine("所选的文件已被移动至其他地方");
            }
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            //string file = @"c:/ windows/notepad.exe"; 
            psi.Arguments = " /select," + filePath;
            System.Diagnostics.Process.Start(psi);
        }
        /// <summary>
        /// 已选择文件TabButton点击事件(添加水印)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabButton_SeletedFile_Click(object sender, RoutedEventArgs e)
        {
            addWaterMarkButton.Visibility = Visibility.Visible;
            SerchFile_Grid.Visibility = Visibility.Collapsed;
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
            this.Dispatcher.Invoke(new Action(() =>
            {
                tabButton_SelectedFiles.BorderBrush = tabButton_BorderBrush_Color_Seleted;
                tabButton_AddedWaterMarkFiles.BorderBrush = tabButton_BorderBrush_Color_Default;
                tabButton_SelectedFiles.Foreground = tabButton_Foreground_Color_Seleted;
                tabButton_AddedWaterMarkFiles.Foreground = tabButton_Foreground_Color_Default;
            })
            );
        }
        /// <summary>
        /// 已添加水印文件夹TabButton点击事件(添加水印)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabButton_AddedWaterMarkFile_Click(object sender, RoutedEventArgs e)
        {
            AddedWatermarkFile_Grid.ItemsSource = null;
            addWaterMarkButton.Visibility = Visibility.Collapsed;
            SerchFile_Grid.Visibility = Visibility.Visible;
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
            this.Dispatcher.Invoke(new Action(() =>
            {
                tabButton_SelectedFiles.BorderBrush = tabButton_BorderBrush_Color_Default;
                tabButton_AddedWaterMarkFiles.BorderBrush = tabButton_BorderBrush_Color_Seleted;
                tabButton_SelectedFiles.Foreground = tabButton_Foreground_Color_Default;
                tabButton_AddedWaterMarkFiles.Foreground = tabButton_Foreground_Color_Seleted;
            })
            );
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
        /// <summary>
        /// 底部标签页向左切换按钮(暂时用不到)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PageUpButton_Click(object sender, RoutedEventArgs e)
        {
            
        }
        /// <summary>
        /// 底部标签页向右切换按钮(暂时用不到)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PageDownButton_Click(object sender, RoutedEventArgs e)
        {
            
        }
        /// <summary>
        /// 展示文件正在添加水印的处理动效
        /// </summary>
        private void ShowProgressBar(int GridIndex)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                FrameworkElement element = templeColumn.GetCellContent(membersDataGrid.Items[GridIndex]);
                if (element != null)
                {
                    System.Windows.Controls.Button removeFileButton = (System.Windows.Controls.Button)templeColumn.CellTemplate.FindName("RemoveFile_Button", element);
                    LoadingCircle loadingCircle = (LoadingCircle)templeColumn.CellTemplate.FindName("FileLoadingCircle", element);
                    System.Windows.Controls.Button openFileButton = (System.Windows.Controls.Button)templeColumn.CellTemplate.FindName("OpenFileButton", element);
                    if (removeFileButton != null)
                    {
                        removeFileButton.Visibility = Visibility.Collapsed;
                        loadingCircle.Visibility = Visibility.Visible;
                        openFileButton.Visibility = Visibility.Collapsed;
                    }
                }
            })
            );
        }
        /// <summary>
        /// 水印添加完毕后，将移除文件按钮替换为打开文件按钮
        /// </summary>
        private void ShowOpenFileButton(int GridIndex)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                FrameworkElement element = templeColumn.GetCellContent(membersDataGrid.Items[GridIndex]);
                if (element != null)
                {
                    System.Windows.Controls.Button removeFileButton = (System.Windows.Controls.Button)templeColumn.CellTemplate.FindName("RemoveFile_Button", element);
                    MahApps.Metro.IconPacks.PackIconMaterial checkIcon = (MahApps.Metro.IconPacks.PackIconMaterial)templeColumn.CellTemplate.FindName("Check_Icon", element);
                    System.Windows.Controls.TextBox addedWaterMark_TextBox = (System.Windows.Controls.TextBox)templeColumn.CellTemplate.FindName("AddedWaterMark_TextBox", element);
                    LoadingCircle loadingCircle = (LoadingCircle)templeColumn.CellTemplate.FindName("FileLoadingCircle", element);
                    System.Windows.Controls.Button openFileButton = (System.Windows.Controls.Button)templeColumn.CellTemplate.FindName("OpenFileButton", element);
                    if (removeFileButton != null)
                    {
                        this.Dispatcher.Invoke(new Action(() =>
                        {
                            removeFileButton.Visibility = Visibility.Collapsed;
                            loadingCircle.Visibility = Visibility.Collapsed;
                            checkIcon.Visibility = Visibility.Visible;
                            addedWaterMark_TextBox.Visibility = Visibility.Visible;
                            openFileButton.Visibility = Visibility.Visible;
                        }));
                    }
                }
            })
            );

        }
        /// <summary>
        /// 展开或收起正在添加水印的提示遮罩
        /// </summary>
        /// <param name="needShow"></param>
        private void ShowAddingWaterMarkMask(bool needShow)
        {
            int RightMargin = 0;
            if (isClose)
            {
                RightMargin = 500;
            }
            else
            {
                RightMargin = 300;
            }
            if (needShow)
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    ThicknessAnimation marginAnimation = new ThicknessAnimation();
                    CubicEase ease = new CubicEase { EasingMode = EasingMode.EaseInOut };
                    marginAnimation.EasingFunction = ease;
                    marginAnimation.From = new Thickness(0, 0, 1080, 20);
                    marginAnimation.To = new Thickness(0, 0, RightMargin, 20);
                    marginAnimation.Duration = TimeSpan.FromSeconds(0.8);
                    AddingWaterMark_Mask.BeginAnimation(Border.MarginProperty, marginAnimation);
                })
                );
            }
            else
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    ThicknessAnimation marginAnimation = new ThicknessAnimation();
                    CubicEase ease = new CubicEase { EasingMode = EasingMode.EaseInOut };
                    marginAnimation.EasingFunction = ease;
                    marginAnimation.From = new Thickness(0, 0, RightMargin, 20);
                    marginAnimation.To = new Thickness(0, 0, 1080, 20);
                    marginAnimation.Duration = TimeSpan.FromSeconds(0.8);
                    AddingWaterMark_Mask.BeginAnimation(Border.MarginProperty, marginAnimation);
                })
                );
            }
        }
        /// <summary>
        /// 点击已选择文件列表界面内的已添加水印的打开文件按钮,调用系统窗口定位文件(添加水印)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            int selectedRowIndex = membersDataGrid.SelectedIndex;
            string fileExtension = System.IO.Path.GetExtension(members[selectedRowIndex].FilePath);
            string filePath = members[selectedRowIndex].FileDir + "\\" + members[selectedRowIndex].FileName + "(已添加水印)" + fileExtension;
            if (!System.IO.File.Exists(filePath))
            {
                //Trace.WriteLine("所选的文件已被移动至其他地方");
            }
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            psi.Arguments = " /select," + filePath;
            System.Diagnostics.Process.Start(psi);
        }
        /// <summary>
        /// 点击帮助Help按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            
        }
        
        /// <summary>
        /// 点击关闭按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CloseTheAppButton_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }
        /// <summary>
        /// 大小写转换
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void a2A_TextBox_a_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (a2A_textBox_a.Text != "")
            {
                
                this.Dispatcher.Invoke(new Action(() =>
                {
                    a2A_Icon.Foreground = a2AIcon_bgColor_Truning;
                    TimeDelay.Delay(50);
                    a2A_Icon.Foreground = a2AIcon_bgColor_Default;
                }));
                a2A_textBox_A.Text = AaConvert.a2Afunc(a2A_textBox_a.Text);
            }
            else
            {
                a2A_textBox_A.Text = "";
            }
        }
        /// <summary>
        /// 将转换后的大写人民币数值拷贝到剪贴板中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Copy_A_Value_Button_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Clipboard.SetDataObject(a2A_textBox_A.Text);
        }
        
        /// <summary>
        /// 点击开始生成报账文档
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MoneyRequestGrid_Start_Button_Click(object sender, RoutedEventArgs e)
        {
            List<string> list_InfoText = new List<string>();
            List<string> list_CostText = new List<string>();
            System.Windows.Controls.TextBox lT_Haocai_Info_TextBox = (System.Windows.Controls.TextBox)LT_Haocai_Info_Grid.FindName("LT_Haocai_Info_TextBox");
            System.Windows.Controls.TextBox lT_Haocai_Cost_TextBox = (System.Windows.Controls.TextBox)LT_Haocai_Cost_Grid.FindName("LT_Haocai_Cost_TextBox");
            System.Windows.Controls.TextBox lT_Weixiu_Info_TextBox = (System.Windows.Controls.TextBox)LT_Weixiu_Info_Grid.FindName("LT_Weixiu_Info_TextBox");
            System.Windows.Controls.TextBox lT_Weixiu_Cost_TextBox = (System.Windows.Controls.TextBox)LT_Weixiu_Cost_Grid.FindName("LT_Weixiu_Cost_TextBox");
            System.Windows.Controls.TextBox hN_Haocai_Info_TextBox = (System.Windows.Controls.TextBox)HN_Haocai_Info_Grid.FindName("HN_Haocai_Info_TextBox");
            System.Windows.Controls.TextBox hN_Haocai_Cost_TextBox = (System.Windows.Controls.TextBox)HN_Haocai_Cost_Grid.FindName("HN_Haocai_Cost_TextBox");
            System.Windows.Controls.TextBox hN_Weixiu_Info_TextBox = (System.Windows.Controls.TextBox)HN_Weixiu_Info_Grid.FindName("HN_Weixiu_Info_TextBox");
            System.Windows.Controls.TextBox hN_Weixiu_Cost_TextBox = (System.Windows.Controls.TextBox)HN_Weixiu_Cost_Grid.FindName("HN_Weixiu_Cost_TextBox");
            if (string.IsNullOrEmpty(lT_Haocai_Info_TextBox.Text))
            {
                list_InfoText.Add("0");
            }
            else
            {
                list_InfoText.Add(lT_Haocai_Info_TextBox.Text);
            }

            if (string.IsNullOrEmpty(lT_Weixiu_Info_TextBox.Text))
            {
                list_InfoText.Add("0");
            }
            else
            {
                list_InfoText.Add(lT_Weixiu_Info_TextBox.Text);
            }

            if (string.IsNullOrEmpty(hN_Haocai_Info_TextBox.Text))
            {
                list_InfoText.Add("0");
            }
            else
            {
                list_InfoText.Add(hN_Haocai_Info_TextBox.Text);
            }

            if (string.IsNullOrEmpty(hN_Weixiu_Info_TextBox.Text))
            {
                list_InfoText.Add("0");
            }
            else
            {
                list_InfoText.Add(hN_Weixiu_Info_TextBox.Text);
            }


            if (string.IsNullOrEmpty(lT_Haocai_Cost_TextBox.Text))
            {
                list_CostText.Add("0");
            }
            else
            {
                list_CostText.Add(lT_Haocai_Cost_TextBox.Text);
            }

            if (string.IsNullOrEmpty(lT_Weixiu_Cost_TextBox.Text))
            {
                list_CostText.Add("0");
            }
            else
            {
                list_CostText.Add(lT_Weixiu_Cost_TextBox.Text);
            }

            if (string.IsNullOrEmpty(hN_Haocai_Cost_TextBox.Text))
            {
                list_CostText.Add("0");
            }
            else
            {
                list_CostText.Add(hN_Haocai_Cost_TextBox.Text);
            }

            if (string.IsNullOrEmpty(hN_Weixiu_Cost_TextBox.Text))
            {
                list_CostText.Add("0");
            }
            else
            {
                list_CostText.Add(hN_Weixiu_Cost_TextBox.Text);
            }

            //Trace.WriteLine("list_InfoText的长度是:" + list_InfoText.Count);
            //Trace.WriteLine("list_CostText的长度是:" + list_CostText.Count);
            for (int i = 0; i < 2; i++)
            {
                //if (list_InfoText[i] != "0")
                //{
                //    fileOperate.InitMoneyRequestDOC(i, list_InfoText[i], list_CostText[i]);
                //}
                //else
                //{
                //    Trace.WriteLine("list_InfoText[" + i + "] = 0");
                //}
            }
            ////撰写联拓转账请示文档
            //fileOperate.InitMoneyRequestDOC(0, list_InfoText[0], list_InfoText[1], list_CostText[0], list_CostText[1]);
            ////撰写海纳转账请示文档
            //fileOperate.InitMoneyRequestDOC(1, list_InfoText[2], list_InfoText[3], list_CostText[2], list_CostText[3]);
            //撰写联拓转账请示文档
            fileOperate.InitMoneyRequestXlsx(0, list_InfoText[0], list_InfoText[1], list_CostText[0], list_CostText[1]);
            //撰写海纳转账请示文档
            fileOperate.InitMoneyRequestXlsx(1, list_InfoText[2], list_InfoText[3], list_CostText[2], list_CostText[3]);
        }


        /// <summary>
        /// 图片裁切页面内顶部标签页按钮(暂时用不到)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureCropping_TabButton_SeletedFile_Click(object sender, RoutedEventArgs e)
        {
            //暂时用不到
        }

        private void PictureCropping_Button_AddFile_Click(object sender, RoutedEventArgs e)
        {
            FillGridData_PictureCropping();
            //button_SelectedFiles_PictureCropping.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Button.ClickEvent));
        }
        /// <summary>
        /// 填充已选择图片列表(图片裁切)
        /// </summary>
        private void FillGridData_PictureCropping()
        {
            //打开系统窗口获取文件路径列表，相同的文件则忽略
            file_list = getFileInfo.GetFilePath();
            int CurrentFileCount = members_PictureCropping.Count;
            for (int i = 0; i < file_list.Count; i++)
            {
                bool flag = true;

                foreach (var file in members_PictureCropping)
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
                    members_PictureCropping.Add(new Member
                    {
                        FilePath = fileFullInfo["filePath"].ToString(),
                        Number = (CurrentFileCount + i + 1).ToString(),
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

            PictureCropping_membersDataGrid.ItemsSource = members_PictureCropping;
            if (members_PictureCropping.Count != 0)
            {
                show_NoFile_Text(false);
            }
            else
            {
                show_NoFile_Text(true);
            }
        }
        /// <summary>
        /// 点击此按钮开始批量裁切图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureCropping_Button_Start_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(cuttingParameters_TextBox.Text))
            {
                string[] CuttingParameters = fileOperate.ProcessString(cuttingParameters_TextBox.Text);
                int leftOffset = int.Parse(CuttingParameters[0]);
                int topOffset = int.Parse(CuttingParameters[1]);
                int rightOffset = int.Parse(CuttingParameters[2]);
                int bottomOffset = int.Parse(CuttingParameters[3]);
                int cornerRadius = int.Parse(CuttingParameters[4]);

                foreach (var file in members_PictureCropping)
                {
                    string filePath = file.FilePath;
                    string fileDir = file.FileDir;
                    fileOperate.StartPictureCropping(filePath, fileDir, leftOffset, rightOffset, topOffset, bottomOffset, cornerRadius);
                }

                Trace.WriteLine("所有图片已处理完成！");
            }
            else 
            {
                int leftOffset = 40;
                int topOffset = 56;
                int rightOffset = 40;
                int bottomOffset = 118;
                int cornerRadius = 20;

                foreach (var file in members_PictureCropping)
                {
                    string filePath = file.FilePath;
                    string fileDir = file.FileDir;
                    fileOperate.StartPictureCropping(filePath, fileDir, leftOffset, rightOffset, topOffset, bottomOffset, cornerRadius);
                }

                Trace.WriteLine("所有图片已处理完成！");
            }
        }

        /// <summary>
        /// 填充已选择Excel列表(微信太阳码)
        /// </summary>
        private void FillGridData_NormalQRCode()
        {
            //打开系统窗口获取文件路径列表，相同的文件则忽略
            file_list = getFileInfo.GetFilePath();
            int CurrentFileCount = members_NormalQRCode.Count;
            for (int i = 0; i < file_list.Count; i++)
            {
                bool flag = true;

                foreach (var file in members_NormalQRCode)
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
                    members_NormalQRCode.Add(new Member
                    {
                        FilePath = fileFullInfo["filePath"].ToString(),
                        Number = (CurrentFileCount + i + 1).ToString(),
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

            NormalQRCode_membersDataGrid.ItemsSource = members_NormalQRCode;
            if (members_NormalQRCode.Count != 0)
            {
                show_NoFile_Text(false);
            }
            else
            {
                show_NoFile_Text(true);
            }
        }
        /// <summary>
        /// 通用二维码添加文件按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NormalQRCode_Button_AddFile_Click(object sender, RoutedEventArgs e)
        {
            FillGridData_NormalQRCode();
        }
        private void NormalQRCode_TabButton_SeletedFile_Click(object sender, RoutedEventArgs e)
        {
            //暂时用不到
        }
        /// <summary>
        /// 点击按钮生成通用二维码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void NormalQRCode_Start_Generate_Click(object sender, RoutedEventArgs e)
        {
            string saveName = "";
            if (NormalQRCode_PagePath_TextBox != null && !string.IsNullOrEmpty(NormalQRCode_PagePath_TextBox.Text))
            {
                if (NormalQRCode_SaveName_TextBox != null && string.IsNullOrEmpty(NormalQRCode_SaveName_TextBox.Text))
                {
                    saveName = "未命名二维码";
                }
                else
                {
                    saveName = NormalQRCode_SaveName_TextBox.Text;
                }
                await NormalQRCodeGenerated.NormalQRCode_Generate(NormalQRCode_PagePath_TextBox.Text, Path.Combine(@"C:\Users\12040\Desktop\二维码", $"{saveName}.png"));
            }


            foreach (var file in members_NormalQRCode)
            {
                string filePath = file.FilePath;
                await NormalQRCodeGenerated.NormalQRCodeMultiplyGenerated(filePath);
            }
        }
        /// <summary>
        /// 填充已选择Excel列表(微信太阳码)
        /// </summary>
        private void FillGridData_WXQRCode()
        {
            //打开系统窗口获取文件路径列表，相同的文件则忽略
            file_list = getFileInfo.GetFilePath();
            int CurrentFileCount = members_WXQRCode.Count;
            for (int i = 0; i < file_list.Count; i++)
            {
                bool flag = true;

                foreach (var file in members_WXQRCode)
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
                    members_WXQRCode.Add(new Member
                    {
                        FilePath = fileFullInfo["filePath"].ToString(),
                        Number = (CurrentFileCount + i + 1).ToString(),
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

            WXQRCode_membersDataGrid.ItemsSource = members_WXQRCode;
            if (members_WXQRCode.Count != 0)
            {
                show_NoFile_Text(false);
            }
            else
            {
                show_NoFile_Text(true);
            }
        }
        /// <summary>
        /// 微信太阳码添加文件按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WXQRCode_Button_AddFile_Click(object sender, RoutedEventArgs e)
        {
            FillGridData_WXQRCode();
        }

        private void WXQRCode_TabButton_SeletedFile_Click(object sender, RoutedEventArgs e)
        {
            //暂时用不到
        }
        /// <summary>
        /// 生成单个太阳码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void WXQRCode_Single_Start_Click(object sender, RoutedEventArgs e)
        {
            await WXQRCodeGenerated.WXQRCodeSingleGenerated(textBox_ChanelName.Text, textBox_PagePath.Text, textBox_SaveName.Text);
        }
        /// <summary>
        /// 批量生成太阳码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void WXQRCode_Multiply_Start_Click(object sender, RoutedEventArgs e)
        {
            int CurrentFileCount = members_WXQRCode.Count;
            if (CurrentFileCount != 0)
            {
                foreach (var file in members_WXQRCode)
                {
                    string filePath = file.FilePath;
                    await WXQRCodeGenerated.WXQRCodeMultiplyGenerated(filePath);
                }
            }
            else 
            {
                await WXQRCodeGenerated.WXQRCodeSingleGenerated(textBox_ChanelName.Text, textBox_PagePath.Text, textBox_SaveName.Text);
            }
            
        }
        /// <summary>
        /// 开始部署微信机器人
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TaskSchedulerDeploy_Start_Button_Click(object sender, RoutedEventArgs e)
        {
            string xmlFilePath = @"C:\Users\12040\source\repos\MyApp\MyApp\bin\Release\net8.0-windows\TaskSchedulerDeploy\每天早上7点发送团队预约信息至开放群.xml";
            string taskName = "每天早上7点整发送团队预约数据到开放群";
            // 检查当前用户是否是管理员
            if (!TaskSchedulerDeploy.IsRunAsAdmin())
            {
                // 重新以管理员权限运行
                var exeName = Process.GetCurrentProcess().MainModule.FileName;
                var startInfo = new ProcessStartInfo(exeName)
                {
                    UseShellExecute = true,
                    Verb = "runas" // 提升为管理员权限
                };

                try
                {
                    Process.Start(startInfo);
                }
                catch
                {
                    Console.WriteLine("无法提升为管理员权限，请手动以管理员身份运行程序！");
                }
                return;
            }
            TaskSchedulerDeploy.ImportTaskFromXmlUsingLibrary(xmlFilePath, taskName);
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
                PictureCroppingGrid.Visibility = Visibility.Collapsed;
                a2A_Grid.Visibility = Visibility.Collapsed;
                MoneyRequestGrid.Visibility = Visibility.Collapsed;
                NormalQRCodeGrid.Visibility = Visibility.Collapsed;
                WXQRCodeGrid.Visibility = Visibility.Collapsed;
                TaskSchedulerDeployGrid.Visibility = Visibility.Collapsed;
                //背景图片路径
                string ImagePath = @"Images/BG7.jpg";
                string absolutePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ImagePath);
                // 切换图片
                CurrentImageBrush.ImageSource = new BitmapImage(new System.Uri(absolutePath, System.UriKind.Relative));
                //显示添加水印界面
                MainGrid.Visibility = Visibility.Visible;
            }));
        }
        /// <summary>
        /// 左侧菜单栏:图片裁切按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_PictureCropping_Click(object sender, RoutedEventArgs e)
        {
            MainGrid.Visibility = Visibility.Collapsed;
            a2A_Grid.Visibility = Visibility.Collapsed;
            MoneyRequestGrid.Visibility = Visibility.Collapsed;
            NormalQRCodeGrid.Visibility = Visibility.Collapsed;
            WXQRCodeGrid.Visibility = Visibility.Collapsed;
            TaskSchedulerDeployGrid.Visibility = Visibility.Collapsed;

            //背景图片路径
            string ImagePath = @"Images/BG7.jpg";
            string absolutePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ImagePath);
            // 切换图片
            CurrentImageBrush.ImageSource = new BitmapImage(new System.Uri(absolutePath, System.UriKind.Relative));
            //显示图片裁切界面
            PictureCroppingGrid.Visibility = Visibility.Visible;
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
                MainGrid.Visibility = Visibility.Collapsed;
                PictureCroppingGrid.Visibility = Visibility.Collapsed;
                MoneyRequestGrid.Visibility = Visibility.Collapsed;
                NormalQRCodeGrid.Visibility = Visibility.Collapsed;
                WXQRCodeGrid.Visibility = Visibility.Collapsed;
                TaskSchedulerDeployGrid.Visibility = Visibility.Collapsed;
                show_NoFile_Text(false);

                //背景图片路径
                string ImagePath = @"Images\BG4.jpg";
                string absolutePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ImagePath);
                // 切换图片
                CurrentImageBrush.ImageSource = new BitmapImage(new System.Uri(absolutePath, System.UriKind.Relative));

                a2A_Grid.Visibility = Visibility.Visible;
            }));
        }
        /// <summary>
        /// 菜单栏：点击撰写转账请示文档按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_MoneyRequest_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                MainGrid.Visibility = Visibility.Collapsed;
                PictureCroppingGrid.Visibility = Visibility.Collapsed;
                a2A_Grid.Visibility = Visibility.Collapsed;
                NormalQRCodeGrid.Visibility = Visibility.Collapsed;
                WXQRCodeGrid.Visibility = Visibility.Collapsed;
                TaskSchedulerDeployGrid.Visibility = Visibility.Collapsed;
                show_NoFile_Text(false);

                //背景图片路径
                string ImagePath = @"Images\BG4.jpg";
                string absolutePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ImagePath);
                // 切换图片
                CurrentImageBrush.ImageSource = new BitmapImage(new System.Uri(absolutePath, System.UriKind.Relative));
                //显示撰写转账请示文档界面
                MoneyRequestGrid.Visibility = Visibility.Visible;
            }));
        }
        /// <summary>
        /// 菜单栏：点击生成通用二维码按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_NormalQRCodeGenerated_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                MainGrid.Visibility = Visibility.Collapsed;
                PictureCroppingGrid.Visibility = Visibility.Collapsed;
                a2A_Grid.Visibility = Visibility.Collapsed;
                MoneyRequestGrid.Visibility = Visibility.Collapsed;
                WXQRCodeGrid.Visibility = Visibility.Collapsed;
                TaskSchedulerDeployGrid.Visibility = Visibility.Collapsed;
                show_NoFile_Text(false);
                //背景图片路径
                string ImagePath = @"Images\BG5.jpg";
                string absolutePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ImagePath);
                // 切换图片
                CurrentImageBrush.ImageSource = new BitmapImage(new System.Uri(absolutePath, System.UriKind.Relative));
                //显示通用二维码生成界面
                NormalQRCodeGrid.Visibility = Visibility.Visible;
            }));
        }
        /// <summary>
        /// 菜单栏按钮：点击微信二维码按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_WXQRCodeGenerated_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                MainGrid.Visibility = Visibility.Collapsed;
                PictureCroppingGrid.Visibility = Visibility.Collapsed;
                a2A_Grid.Visibility = Visibility.Collapsed;
                MoneyRequestGrid.Visibility = Visibility.Collapsed;
                NormalQRCodeGrid.Visibility = Visibility.Collapsed;
                TaskSchedulerDeployGrid.Visibility = Visibility.Collapsed;
                show_NoFile_Text(false);
                //背景图片路径
                string ImagePath = @"Images\BG5.jpg";
                string absolutePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ImagePath);
                // 切换图片
                CurrentImageBrush.ImageSource = new BitmapImage(new System.Uri(absolutePath, System.UriKind.Relative));
                //显示太阳码生成器界面
                WXQRCodeGrid.Visibility = Visibility.Visible;
            }));
        }
        /// <summary>
        /// 菜单栏按钮：点击团队预约自动化按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuButton_WXBotDeploy_Click(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                MainGrid.Visibility = Visibility.Collapsed;
                PictureCroppingGrid.Visibility = Visibility.Collapsed;
                a2A_Grid.Visibility = Visibility.Collapsed;
                MoneyRequestGrid.Visibility = Visibility.Collapsed;
                NormalQRCodeGrid.Visibility = Visibility.Collapsed;
                WXQRCodeGrid.Visibility = Visibility.Collapsed;
                show_NoFile_Text(false);
                //背景图片路径
                string ImagePath = @"Images\BG7.jpg";
                string absolutePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ImagePath);
                // 切换图片
                CurrentImageBrush.ImageSource = new BitmapImage(new System.Uri(absolutePath, System.UriKind.Relative));
                //显示团队预约自动化界面
                TaskSchedulerDeployGrid.Visibility = Visibility.Visible;
            }));
        }
        /// <summary>
        /// 点击打开或关闭菜单按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        bool isClose = true;//菜单栏是否折叠
        private void OpenOrCLoseMenuButton_Click(object sender, RoutedEventArgs e)
        {
            MenuGrid_Width.Width = new GridLength(200);
            if (isClose)
            {
                Trace.WriteLine("——————————————菜单栏已折叠————————————————");
                this.Dispatcher.Invoke(new Action(() =>
                {
                    this.IsEnabled = false;
                    Storyboard storyboard = new Storyboard();
                    Duration duration = new Duration(TimeSpan.FromMilliseconds(500));
                    CubicEase ease = new CubicEase { EasingMode = EasingMode.EaseInOut };
                    DoubleAnimation animation = new DoubleAnimation();
                    DoubleAnimation animation_Width = new DoubleAnimation();
                    DoubleAnimation animation_Height = new DoubleAnimation();
                    DoubleAnimation animation_Image_LilyOfTheValley_Width = new DoubleAnimation();
                    DoubleAnimation animation_Image_MorningGlory_Width = new DoubleAnimation();

                    animation.EasingFunction = ease;
                    animation.Duration = duration;
                    storyboard.Children.Add(animation);
                    animation.From = 0;
                    animation.To = 200;

                    animation_Width.EasingFunction = ease;
                    animation_Width.Duration = duration;
                    storyboard.Children.Add(animation_Width);
                    animation_Width.From = 0;
                    animation_Width.To = 80;

                    animation_Height.EasingFunction = ease;
                    animation_Height.Duration = duration;
                    storyboard.Children.Add(animation_Height);
                    animation_Height.From = 0;
                    animation_Height.To = 80;

                    animation_Image_LilyOfTheValley_Width.EasingFunction = ease;
                    animation_Image_LilyOfTheValley_Width.Duration = duration;
                    storyboard.Children.Add(animation_Image_LilyOfTheValley_Width);
                    animation_Image_LilyOfTheValley_Width.From = 394;
                    animation_Image_LilyOfTheValley_Width.To = 0;

                    animation_Image_MorningGlory_Width.EasingFunction = ease;
                    animation_Image_MorningGlory_Width.Duration = duration;
                    storyboard.Children.Add(animation_Image_MorningGlory_Width);
                    animation_Image_MorningGlory_Width.From = 0;
                    animation_Image_MorningGlory_Width.To = 196;

                    Storyboard.SetTarget(animation, MenuGrid_Width);
                    Storyboard.SetTarget(animation_Width, p1_Image_Border);
                    Storyboard.SetTarget(animation_Height, p1_Image_Border);
                    Storyboard.SetTarget(animation_Image_LilyOfTheValley_Width, Image_LilyOfTheValley);
                    Storyboard.SetTarget(animation_Image_MorningGlory_Width, Image_MorningGlory);

                    Storyboard.SetTargetProperty(animation, new PropertyPath("(ColumnDefinition.MaxWidth)"));
                    Storyboard.SetTargetProperty(animation_Width, new PropertyPath("Width"));
                    Storyboard.SetTargetProperty(animation_Height, new PropertyPath("Height"));
                    Storyboard.SetTargetProperty(animation_Image_LilyOfTheValley_Width, new PropertyPath("Width"));
                    Storyboard.SetTargetProperty(animation_Image_MorningGlory_Width, new PropertyPath("Width"));

                    storyboard.Begin();

                    this.IsEnabled = true;
                    openOrCLoseMenuButton_Icon.Kind = MahApps.Metro.IconPacks.PackIconMaterialKind.MenuLeft;
                })
                );
            }
            else
            {
                Trace.WriteLine("——————————————菜单栏已展开————————————————");
                this.Dispatcher.Invoke(new Action(() =>
                {
                    this.IsEnabled = false;
                    Storyboard storyboard = new Storyboard();
                    Duration duration = new Duration(TimeSpan.FromMilliseconds(500));
                    CubicEase ease = new CubicEase { EasingMode = EasingMode.EaseInOut };
                    DoubleAnimation animation = new DoubleAnimation();
                    DoubleAnimation animation_Width = new DoubleAnimation();
                    DoubleAnimation animation_Height = new DoubleAnimation();
                    DoubleAnimation animation_Image_LilyOfTheValley_Width = new DoubleAnimation();
                    DoubleAnimation animation_Image_MorningGlory_Width = new DoubleAnimation();

                    animation.EasingFunction = ease;
                    animation.Duration = duration;
                    storyboard.Children.Add(animation);
                    animation.From = 200;
                    animation.To = 0;

                    animation_Width.EasingFunction = ease;
                    animation_Width.Duration = duration;
                    storyboard.Children.Add(animation_Width);
                    animation_Width.From = 80;
                    animation_Width.To = 0;

                    animation_Height.EasingFunction = ease;
                    animation_Height.Duration = duration;
                    storyboard.Children.Add(animation_Height);
                    animation_Height.From = 80;
                    animation_Height.To = 0;

                    animation_Image_LilyOfTheValley_Width.EasingFunction = ease;
                    animation_Image_LilyOfTheValley_Width.Duration = duration;
                    storyboard.Children.Add(animation_Image_LilyOfTheValley_Width);
                    animation_Image_LilyOfTheValley_Width.From = 0;
                    animation_Image_LilyOfTheValley_Width.To = 394;

                    animation_Image_MorningGlory_Width.EasingFunction = ease;
                    animation_Image_MorningGlory_Width.Duration = duration;
                    storyboard.Children.Add(animation_Image_MorningGlory_Width);
                    animation_Image_MorningGlory_Width.From = 196;
                    animation_Image_MorningGlory_Width.To = 0;

                    Storyboard.SetTarget(animation, MenuGrid_Width);
                    Storyboard.SetTarget(animation_Width, p1_Image_Border);
                    Storyboard.SetTarget(animation_Height, p1_Image_Border);
                    Storyboard.SetTarget(animation_Image_LilyOfTheValley_Width, Image_LilyOfTheValley);
                    Storyboard.SetTarget(animation_Image_MorningGlory_Width, Image_MorningGlory);

                    Storyboard.SetTargetProperty(animation, new PropertyPath("(ColumnDefinition.MaxWidth)"));
                    Storyboard.SetTargetProperty(animation_Width, new PropertyPath("Width"));
                    Storyboard.SetTargetProperty(animation_Height, new PropertyPath("Height"));
                    Storyboard.SetTargetProperty(animation_Image_LilyOfTheValley_Width, new PropertyPath("Width"));
                    Storyboard.SetTargetProperty(animation_Image_MorningGlory_Width, new PropertyPath("Width"));

                    storyboard.Begin();

                    this.IsEnabled = true;
                    openOrCLoseMenuButton_Icon.Kind = MahApps.Metro.IconPacks.PackIconMaterialKind.MenuRight;
                })
                );
            }
            isClose = !isClose;
        }
        /// <summary>
        /// 是否是AR导览内的链接：是
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_IsAR_Checked(object sender, RoutedEventArgs e)
        {
            Thickness oldThickness = new Thickness(50, 0, 200, 0);
            Thickness newThickness = new Thickness(250, 0, 200, 0);
            
            this.Dispatcher.Invoke(new Action(() =>
            {
                Storyboard storyboard = new Storyboard();
                Duration duration = new Duration(TimeSpan.FromMilliseconds(500));
                CubicEase ease = new CubicEase { EasingMode = EasingMode.EaseInOut };
                DoubleAnimation animation_Width = new DoubleAnimation();
                ThicknessAnimation animation_Margin = new ThicknessAnimation();

                animation_Width.EasingFunction = ease;
                animation_Width.Duration = duration;
                storyboard.Children.Add(animation_Width);
                animation_Width.From = 0;
                animation_Width.To = 200;

                animation_Margin.EasingFunction = ease;
                animation_Margin.Duration = duration;
                storyboard.Children.Add(animation_Margin);
                animation_Margin.From = oldThickness;
                animation_Margin.To = newThickness;

                Storyboard.SetTarget(animation_Width, wXQRCode_ChanelName_TextBox_Grid);
                Storyboard.SetTarget(animation_Margin, WXQRCode_PagePath_TextBox_Grid);

                Storyboard.SetTargetProperty(animation_Width, new PropertyPath("Width"));
                Storyboard.SetTargetProperty(animation_Margin, new PropertyPath("Margin"));

                storyboard.Begin();
            })
            );
        }
        /// <summary>
        /// 是否是AR导览内的链接：否
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_IsAR_Unchecked(object sender, RoutedEventArgs e)
        {
            Thickness oldThickness = new Thickness(250, 0, 200, 0);
            Thickness newThickness = new Thickness(50, 0, 200, 0);
            this.Dispatcher.Invoke(new Action(() =>
            {
                Storyboard storyboard = new Storyboard();
                Duration duration = new Duration(TimeSpan.FromMilliseconds(500));
                CubicEase ease = new CubicEase { EasingMode = EasingMode.EaseInOut };
                DoubleAnimation animation_Width = new DoubleAnimation();
                ThicknessAnimation animation_Margin = new ThicknessAnimation();

                animation_Width.EasingFunction = ease;
                animation_Width.Duration = duration;
                storyboard.Children.Add(animation_Width);
                animation_Width.From = 200;
                animation_Width.To = 0;

                animation_Margin.EasingFunction = ease;
                animation_Margin.Duration = duration;
                storyboard.Children.Add(animation_Margin);
                animation_Margin.From = oldThickness;
                animation_Margin.To = newThickness;

                Storyboard.SetTarget(animation_Width, wXQRCode_ChanelName_TextBox_Grid);
                Storyboard.SetTarget(animation_Margin, WXQRCode_PagePath_TextBox_Grid);

                Storyboard.SetTargetProperty(animation_Width, new PropertyPath("Width"));
                Storyboard.SetTargetProperty(animation_Margin, new PropertyPath("Margin"));

                storyboard.Begin();
            })
            );
        }

        private void Button_Test_Click(object sender, RoutedEventArgs e)
        {
            fileOperate.PdfCompress(@"C:\Users\12040\Desktop\测试3.pdf", @"C:\Users\12040\Desktop\水印工具Output\PDF压缩输出文件夹\输出.pdf",10);
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