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
using HandyControl.Controls;
using System.Threading;
using static System.Resources.ResXFileRef;
using System.Windows.Media.Animation;

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
        System.Windows.Controls.Button tabButton_SelectedFiles;
        System.Windows.Controls.Button tabButton_AddedWaterMarkFiles;
        System.Windows.Media.Brush tabButton_BorderBrush_Color_Default;
        System.Windows.Media.Brush tabButton_BorderBrush_Color_Seleted;
        System.Windows.Media.Brush tabButton_Foreground_Color_Default;
        System.Windows.Media.Brush tabButton_Foreground_Color_Seleted;
        Border a2a_Panel;
        Border qRCode_Panel;
        System.Windows.Controls.TextBox textBox_a;
        System.Windows.Controls.TextBox textBox_A;
        System.Windows.Controls.Image qRCode_Image;
        System.Windows.Media.BrushConverter converter = new System.Windows.Media.BrushConverter();//改变首字符圈圈颜色用的
        DataGridTemplateColumn templeColumn;//动态生成的列表
        MahApps.Metro.IconPacks.PackIconMaterial _a2A_Icon;
        System.Windows.Media.Brush a2AIcon_bgColor_Default;
        System.Windows.Media.Brush a2AIcon_bgColor_Truning;
        CircleProgressBar myCircleProgressBar;
        MahApps.Metro.IconPacks.PackIconMaterial addingWaterMark_Icon;

        public MainWindow()
        {
            InitializeComponent();
            title_TextBlock = (TextBlock)MainGrid.FindName("Title_TextBlock");
            text_NoFile = (TextBlock)FatherGrid.FindName("Text_NoFiles");
            button_AddFile = (System.Windows.Controls.Button)MainGrid.FindName("Button_AddFile");
            menuButton_AddWaterMark = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_AddWaterMark");
            menuButton_A2a = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_A2a");
            menuButton_QRCodeGenerated = (System.Windows.Controls.Button)MenuButton_Grid.FindName("MenuButton_QRCodeGenerated");
            tabButton_SelectedFiles = (System.Windows.Controls.Button)MenuButton_Grid.FindName("TabButton_SelectedFiles");
            tabButton_AddedWaterMarkFiles = (System.Windows.Controls.Button)MenuButton_Grid.FindName("TabButton_AddedWaterMarkFiles");
            a2a_Panel = (Border)MainGrid.FindName("A2a_Panel");
            qRCode_Panel = (Border)MainGrid.FindName("QRCode_Panel");
            textBox_a = (System.Windows.Controls.TextBox)a2a_Panel.FindName("TextBox_a");
            textBox_A = (System.Windows.Controls.TextBox)a2a_Panel.FindName("TextBox_A");
            qRCode_Image = (System.Windows.Controls.Image)QRCode_Panel.FindName("QRCode_Image");
            file_list.Clear();
            templeColumn = membersDataGrid.Columns[4] as DataGridTemplateColumn;
            _a2A_Icon = (MahApps.Metro.IconPacks.PackIconMaterial)MainGrid.FindName("a2A_Icon");
            a2AIcon_bgColor_Default = (System.Windows.Media.Brush)converter.ConvertFromString("#FFA5A5A5");
            a2AIcon_bgColor_Truning = (System.Windows.Media.Brush)converter.ConvertFromString("#FF6EA1F3");
            tabButton_BorderBrush_Color_Default = (System.Windows.Media.Brush)converter.ConvertFromString("#DAE2EA");
            tabButton_BorderBrush_Color_Seleted = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");
            tabButton_Foreground_Color_Default = (System.Windows.Media.Brush)converter.ConvertFromString("#FF121518");
            tabButton_Foreground_Color_Seleted = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");
            myCircleProgressBar = (CircleProgressBar)MainGrid.FindName("MyCircleProgressBar");
            //myGifImage = (GifImage)MainGrid.FindName("MyGifImage");
            //myGifImage.Uri = new Uri("pack://siteoforigin:,,,/C:\\Git\\WPF\\WPF-DataTable-Dashboard-master\\Images\\3.gif");
            addingWaterMark_Icon = (MahApps.Metro.IconPacks.PackIconMaterial)MainGrid.FindName("AddingWaterMark_Icon");
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
        /// <summary>
        /// 填充已选择文件列表
        /// </summary>
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
        /// 获取文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetFilePathButton_Click(object sender, RoutedEventArgs e)
        {
            FillGridData();
        }
        /// <summary>
        /// 点击开始添加水印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddWaterMarkButton_Click(object sender, RoutedEventArgs e)
        {
            Border addingWaterMark_Mask = (Border)MainGrid.FindName("AddingWaterMark_Mask");
            System.Windows.Controls.TextBox addingWaterMark_TextBox = (System.Windows.Controls.TextBox)AddingWaterMark_Mask.FindName("AddingWaterMark_TextBox");
            //MahApps.Metro.IconPacks.PackIconMaterial addingWaterMark_Icon = (MahApps.Metro.IconPacks.PackIconMaterial)MainGrid.FindName("AddingWaterMark_Icon");
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
                        //Trace.WriteLine("members[" + i + "].Flag:" + members[i].Flag);
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
                        TimeDelay.Delay(300);
                        addingWaterMark_Mask.Visibility = Visibility.Collapsed;
                        addingWaterMark_Icon.Visibility = Visibility.Collapsed;
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
                //Trace.WriteLine("所选的文件已被移动至其他地方");
            }
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            //string file = @"c:/ windows/notepad.exe"; 
            psi.Arguments = " /select," + filePath;
            System.Diagnostics.Process.Start(psi);
        }
        /// <summary>
        /// 已选择文件TabButton点击事件
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
        private void PageUpButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("已经是第一页");
        }
        private void PageDownButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("已经是最后一页");
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
            
            if (needShow)
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    ThicknessAnimation marginAnimation = new ThicknessAnimation();
                    marginAnimation.From = new Thickness(0, 0, 880, 20);
                    marginAnimation.To = new Thickness(0, 0, 290, 20);
                    marginAnimation.Duration = TimeSpan.FromSeconds(0.3);
                    AddingWaterMark_Mask.BeginAnimation(Border.MarginProperty, marginAnimation);
                })
                );
            }
            else
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    ThicknessAnimation marginAnimation = new ThicknessAnimation();
                    marginAnimation.From = new Thickness(0, 0, 290, 20);
                    marginAnimation.To = new Thickness(0, 0, 880, 20);
                    marginAnimation.Duration = TimeSpan.FromSeconds(0.3);
                    AddingWaterMark_Mask.BeginAnimation(Border.MarginProperty, marginAnimation);
                })
                );
            }
        }
        /// <summary>
        /// 点击已选择文件列表界面内的已添加水印的打开文件按钮,调用系统窗口定位文件
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
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            
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
                this.Dispatcher.Invoke(new Action(() =>
                {
                    _a2A_Icon.Foreground = a2AIcon_bgColor_Truning;
                    TimeDelay.Delay(50);
                    _a2A_Icon.Foreground = a2AIcon_bgColor_Default;
                }));
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