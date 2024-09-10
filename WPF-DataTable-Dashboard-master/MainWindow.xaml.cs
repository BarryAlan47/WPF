﻿using System.Collections.Generic;
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

namespace DataGrid
{
    public partial class MainWindow : Window
    {
        ObservableCollection<Member> members = new ObservableCollection<Member>();
        File file = new File();
        FileOperate fileOperate = new FileOperate();
        List<string> file_list = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
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
        private void GetFilePathButton_Click(object sender, RoutedEventArgs e)
        {
            //获取文件前的fileList的数量
            int BeforeAddFile = file_list.Count;
            //获取选择的文件列表
            file_list = file.GetFilePath();
            //改变首字符圈圈颜色用的
            var converter = new System.Windows.Media.BrushConverter();

            members.Clear();
            membersDataGrid.ItemsSource = null;

            for (int i = 0; i < file_list.Count; i++)
            {
                string fileName = file.getFileName(file_list[i]);
                string fileDir = file.getFileDir(file_list[i]);
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

                members.Add(new Member { Number = (i + 1).ToString(), Character = fileName.Substring(0, 1), BgColor = bgColor, Name = fileName, Position = fileDir, Email = "", Phone = fileType });
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
        private void AddWaterMarkButton_Click(object sender, RoutedEventArgs e) 
        {
            Border addingWaterMark_Mask = (Border)MainGrid.FindName("AddingWaterMark_Mask");

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

                        file_list = file.GetFilePath();

                        var converter = new System.Windows.Media.BrushConverter();//改变首字符圈圈颜色用的

                        members.Clear();
                        membersDataGrid.ItemsSource = null;

                        for (int i = 0; i < file_list.Count; i++)
                        {
                            string fileName = file.getFileName(file_list[i]);
                            string fileDir = file.getFileDir(file_list[i]);
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
                            
                            members.Add(new Member { Number = (i + 1).ToString(), Character = fileName.Substring(0, 1), BgColor = bgColor, Name = fileName, Position = fileDir, Email = "", Phone = fileType });
                        }  
                        membersDataGrid.ItemsSource = members;
                        break;
                }
            }
            else 
            {
                addingWaterMark_Mask.Visibility = Visibility.Visible;
                for (int i = 0; i < file_list.Count; i++)
                {
                    string filePath = file_list[i];
                    string fileDir = file.getFileDir(filePath);
                    string fileName = file.getFileName(filePath);
                    string fileExtension = System.IO.Path.GetExtension(file_list[i]);
                    file.StartAddWaterMark(fileOperate, filePath, fileDir, fileName, fileExtension);

                    //System.Windows.Controls.Button ButtonGroup = (membersDataGrid.Columns[i].GetCellContent(membersDataGrid.Items[5]) as System.Windows.Controls.Button);
                    
                    //StackPanel dataGridTemplateColumn = (StackPanel)membersDataGrid.Columns[i].GetCellContent(membersDataGrid.Items[4]);
                    //System.Windows.Controls.Button removeFile_Button = (System.Windows.Controls.Button)dataGridTemplateColumn.FindName("RemoveFile_Button");
                    //System.Windows.Controls.Button fileCheck_Button = (System.Windows.Controls.Button)dataGridTemplateColumn.FindName("FileCheck_Button");
                    //removeFile_Button.Visibility = System.Windows.Visibility.Collapsed;
                    //fileCheck_Button.Visibility =System.Windows.Visibility.Visible;
                }
                addingWaterMark_Mask.Visibility = Visibility.Collapsed;
            }
        }
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
        private void FileCheck_Button_Click(object sender, RoutedEventArgs e)
        {
            
        }
        //此Boolean值仅供TabButton改变样式使用
        Boolean TabButtonSelect = true;
        //已选择文件夹TabButton点击事件
        private void TabButton_SeletedFile_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button tabButton_SelectedFiles = (System.Windows.Controls.Button)TabButtonGroup.FindName("TabButton_SelectedFiles");
            System.Windows.Controls.Button tabButton_AddedWaterMarkFiles = (System.Windows.Controls.Button)TabButtonGroup.FindName("TabButton_AddedWaterMarkFiles");
            System.Windows.Media.Brush TabButtonBgColor;
            System.Windows.Media.Brush TabButtonForeground;
            System.Windows.Media.Brush TabButtonFontBlack;
            var converter = new System.Windows.Media.BrushConverter();
            TabButtonBgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");
            TabButtonForeground = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");//#FF459ACD
            TabButtonFontBlack = (System.Windows.Media.Brush)converter.ConvertFromString("#FF121518");

            if (TabButtonSelect)
            {
                tabButton_SelectedFiles.BorderBrush = TabButtonBgColor;
                tabButton_SelectedFiles.Foreground = TabButtonForeground;
                tabButton_AddedWaterMarkFiles.BorderBrush = null;
                tabButton_AddedWaterMarkFiles.Foreground = TabButtonFontBlack;
            }
        }
        //已添加水印文件夹TabButton点击事件
        private void TabButton_AddedWaterMarkFile_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button tabButton_SelectedFiles = (System.Windows.Controls.Button)TabButtonGroup.FindName("TabButton_SelectedFiles");
            System.Windows.Controls.Button tabButton_AddedWaterMarkFiles = (System.Windows.Controls.Button)TabButtonGroup.FindName("TabButton_AddedWaterMarkFiles");
            System.Windows.Media.Brush TabButtonBgColor;
            System.Windows.Media.Brush TabButtonForeground;
            System.Windows.Media.Brush TabButtonFontBlack;
            var converter = new System.Windows.Media.BrushConverter();
            TabButtonBgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");
            TabButtonForeground = (System.Windows.Media.Brush)converter.ConvertFromString("#784FF2");
            TabButtonFontBlack = (System.Windows.Media.Brush)converter.ConvertFromString("#FF121518");
            if (TabButtonSelect)
            {
                tabButton_SelectedFiles.BorderBrush = null;
                tabButton_SelectedFiles.Foreground = TabButtonFontBlack;
                
                tabButton_AddedWaterMarkFiles.BorderBrush = TabButtonBgColor;
                tabButton_AddedWaterMarkFiles.Foreground = TabButtonForeground;
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
            fileOperate.PDFSplit();
        }


        private void CloseTheAppButton_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
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
    public class FileOperate
    {
        //加载PDF水印图片
        System.Drawing.Image image = System.Drawing.Image.FromFile(Environment.CurrentDirectory + "\\WaterMarkPic\\PDFWaterMark.png");
        FileStream fs = new System.IO.FileStream(Environment.CurrentDirectory + "\\WaterMarkPic\\PDFWaterMark.png", System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //加载Doc水印图片
        PictureWatermark picture = new PictureWatermark();
        FileStream fileStream = new System.IO.FileStream(Environment.CurrentDirectory + "\\WaterMarkPic\\WordWaterMark.png", System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //加载Excel水印图片
        SkiaSharp.SKBitmap bm = SkiaSharp.SKBitmap.Decode(Environment.CurrentDirectory + "\\WaterMarkPic\\ExcelWaterMark.png");

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
            Worksheet sheet = workbook.Worksheets[0];

            //加载一张图片并设置为背景图片
            //SkiaSharp.SKBitmap bm = SkiaSharp.SKBitmap.Decode(Environment.CurrentDirectory + "\\WaterMarkPic\\ExcelWaterMark.png");
            sheet.PageSetup.BackgoundImage = bm;

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
        public void PDFSplit()
        {
            //SetPDFWatermark("a","b");
        }
    }
    public class File
    {
        public List<string> list = new List<string>();
        //选取文件，并获得路径
        public List<string> GetFilePath()
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog
            {
                Multiselect = true,
                Filter = "Office Files|*.doc;*.docx;*.xls;*.xlsx;*.pdf" //删选、设定文件显示类型
            }; 
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (string filePath in ofd.FileNames)
                {
                    bool flag = true;
                    foreach (var item in list)
                    {
                        if(item == filePath) { flag = false; break; }
                    }
                    if (flag)
                    {
                        list.Add(filePath);
                    }
                }
            }
            return list;
        }
        public string getFileDir(string filePath)
        {
            string fileDir = System.IO.Path.GetDirectoryName(filePath);
            return fileDir;
        }
        public string getFileName(string filePath)
        {
            // 使用Path.GetFileNameWithoutExtension获取不带扩展名的文件名
            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);

            // 使用Path.GetExtension获取文件扩展名（包括点）
            string fileExtension = System.IO.Path.GetExtension(filePath);

            //return fileNameWithoutExt + fileExtension;
            return fileNameWithoutExt;
        }
        public string SetSavePath()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            string savePath = fbd.SelectedPath; //获得选择的文件夹路径
            return savePath;
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
        public string getExecutablePath()
        {
            return Environment.CurrentDirectory;
        }
    }
}