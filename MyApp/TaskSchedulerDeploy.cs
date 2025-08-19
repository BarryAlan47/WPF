using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MyApp
{
    class TaskSchedulerDeploy
    {
        public static bool IsRunAsAdmin()
        {
            using (WindowsIdentity id = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(id);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        public static void ImportTaskFromXmlUsingLibrary(string xmlFilePath, string taskName)
        {
            // 确保路径和任务名称合法
            if (string.IsNullOrWhiteSpace(xmlFilePath) || string.IsNullOrWhiteSpace(taskName))
            {
                throw new ArgumentException("任务名称和XML文件路径不能为空！");
            }

            // 加载 XML 文件内容
            XmlDocument taskXml = new XmlDocument();
            taskXml.Load(xmlFilePath);

            // 连接到本地任务计划程序
            using (TaskService ts = new TaskService())
            {
                // 检查是否已有同名任务存在
                if (ts.GetTask(taskName) != null)
                {
                    throw new InvalidOperationException($"任务 '{taskName}' 已存在！");
                }

                // 从 XML 创建新任务
                ts.RootFolder.ImportTask(taskName, xmlFilePath);
                Console.WriteLine($"任务 '{taskName}' 成功导入！");
            }
        }
    }
}
