using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using static System.Resources.ResXFileRef;

namespace DataGrid
{
    public class GetFileInfo
    {
        public List<string> file_List = new List<string>();
        public Hashtable GetFileFullInfo(string filePath) 
        {
            Hashtable fileFullInfo = new Hashtable();

            fileFullInfo.Clear();

            System.Windows.Media.Brush bgColor;

            var converter = new System.Windows.Media.BrushConverter();

            string fileExtension = System.IO.Path.GetExtension(filePath);

            string fileType;

            if (fileExtension == ".pdf")
            {
                fileType = "PDF文件";
                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#FF5252");
            }
            else if (fileExtension == ".doc" || fileExtension == ".docx")
            {
                fileType = "Word文档";
                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#1E88E5");
            }
            else if (fileExtension == ".xls" || fileExtension == ".xlsx")
            {
                fileType = "Excel表格";
                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#0CA678");
            }
            else
            {
                fileType = "未知类型文件";
                bgColor = (System.Windows.Media.Brush)converter.ConvertFromString("#D3D3D3");
            }
            fileFullInfo.Add("filePath",filePath);
            fileFullInfo.Add("fileDir", GetFileDir(filePath));
            fileFullInfo.Add("fileName", GetFileName(filePath));
            fileFullInfo.Add("fileExtension", fileExtension);
            fileFullInfo.Add("fileType", fileType);
            fileFullInfo.Add("bgColor", bgColor);
            fileFullInfo.Add("addWaterMarkDate", System.DateTime.Now.ToString("d"));

            return fileFullInfo;
        }
        //选取文件，并获得路径
        public List<string> GetFilePath()
        {
            file_List.Clear();
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
                    foreach (var item in file_List)
                    {
                        if (item == filePath) { flag = false; break; }
                    }
                    if (flag)
                    {
                        file_List.Add(filePath);
                    }
                }
            }
            return file_List;
        }
        public string GetFileDir(string filePath)
        {
            string fileDir = System.IO.Path.GetDirectoryName(filePath);
            return fileDir;
        }
        public string GetFileName(string filePath)
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
        public string GetExecutablePath()
        {
            return Environment.CurrentDirectory;
        }
    }
}
