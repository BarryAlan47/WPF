using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace DataGrid
{
    public class GetFileInfo
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
                        if (item == filePath) { flag = false; break; }
                    }
                    if (flag)
                    {
                        list.Add(filePath);
                    }
                }
            }
            return list;
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
