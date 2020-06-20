using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using IWshRuntimeLibrary;
using File = System.IO.File;
using System.Drawing.Printing;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Net.NetworkInformation;

namespace autofix
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            PrinterOperate.GetPrinters(this.label3, this.comboBox2);
            if (NetworkOperate.Ping("192.168.1.3"))
            {
                this.label6.Text = "● 网络连接正常";
                this.label6.ForeColor = System.Drawing.Color.ForestGreen;
                this.label5.Text = "本机IP：" + NetworkOperate.GetLocalIP();
                string strServer = @"\\192.168.1.3";
                string strUserName = "administrator";
                string strUserPD = "ybsnyyzq";
                System.Diagnostics.Process.Start("net.exe", "use   \\\\" + strServer + "     /user:\"" + strUserName + "\"   \"" + strUserPD + "\"");
            }
            else
            {
                this.label6.Text = "● 网络连接异常";
                this.label6.ForeColor = System.Drawing.Color.Crimson;
                MessageBox.Show("内网已断开连接， 请检查网线以及交换机");
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (MessageBox.Show( "选择为：" + this.comboBox1.Text, "确认选择正确？", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string destPath = FilesOperate.GetPathInLnkFile(@"C:\Users\Administrator\Desktop\"+ this.comboBox1.Text + ".lnk");
                BeginWork();
                FilesOperate.MoveFolder(@"\\192.168.1.3\his90程序\"+ this.comboBox1.Text, destPath, this.progressBar1);
                OverWork("修复完成！");
            }
            
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string printer = this.comboBox2.SelectedItem.ToString();
            if (this.comboBox2.SelectedItem != null) //判断是否有选中值
            {

                if (PrinterOperate.SetDefaultPrinter(printer)) //设置默认打印机
                {
                    this.label3.Text = "默认打印机:" + printer;
                    MessageBox.Show(printer + "设置为默认打印机成功！");
                }
                else
                {
                    MessageBox.Show(printer + "设置为默认打印机失败！");
                }
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedText = this.comboBox3.Text;
            string[] exeMap = {"zygl","zyys","ysgl","mzgl","bagl","hqgl","kfgl","yfgl","kfgl","yfgl","zwkf"};

            if (MessageBox.Show("选择为：" + selectedText, "确认选择正确？", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string destPath = @"D:\His程序\" + selectedText;
                BeginWork();
                FilesOperate.MoveFolder(@"\\192.168.1.3\his90程序\" + selectedText, destPath, this.progressBar1);
                FilesOperate.CreateDesktopShortcut(selectedText, destPath + "\\" + exeMap[this.comboBox3.SelectedIndex] + ".exe");
                OverWork("安装完成！");
            }
        }
        private void BeginWork()
        {
            this.progressBar1.Value = 0;
            this.Text = "开始修复.....";
        }
        private void OverWork(string message)
        {
            this.progressBar1.Value = 100;
            MessageBox.Show(message);
            this.progressBar1.Value = 0;
            this.Text = "自助修复工具";
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (NetworkOperate.Ping("192.168.1.3"))
            {
                this.label6.Text = "● 网络连接正常";
                this.label6.ForeColor = System.Drawing.Color.ForestGreen;
                this.label5.Text = "本机IP：" + NetworkOperate.GetLocalIP();
                string strServer = @"\\192.168.1.3";
                string strUserName = "administrator";
                string strUserPD = "ybsnyyzq";
                System.Diagnostics.Process.Start("net.exe", "use   \\\\" + strServer + "     /user:\"" + strUserName + "\"   \"" + strUserPD + "\"");
            }
            else
            {
                this.label6.Text = "● 网络连接异常";
                this.label6.ForeColor = System.Drawing.Color.Crimson;
                MessageBox.Show("内网已断开连接， 请检查网线以及交换机");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }

}






namespace autofix
{
    public class FilesOperate
    {
        public static string GetPathInLnkFile(string path)
        {
            if (File.Exists(path))
            {
                WshShell shell = new WshShell();
                IWshShortcut lnkFile = (IWshShortcut)shell.CreateShortcut(path);
                //快捷方式文件指向的路径.Text = 当前快捷方式文件IWshShortcut类.TargetPath;
                //快捷方式文件指向的目标目录.Text = 当前快捷方式文件IWshShortcut类.WorkingDirectory;
                return lnkFile.WorkingDirectory;
            }
            else
            {
                return "";
            }
        }
        /// <summary>
        /// 移动文件夹中的所有文件夹与文件到另一个文件夹 //转载请注明来自 http://www.uzhanbao.com
        /// </summary>
        /// <param name="sourcePath">源文件夹</param>
        /// <param name="destPath">目标文件夹</param>
        public static void MoveFolder(string sourcePath, string destPath, ProgressBar progressBar)
        {
            if (Directory.Exists(sourcePath))
            {
                if (!Directory.Exists(destPath))
                {
                    //目标目录不存在则创建
                    try
                    {
                        Directory.CreateDirectory(destPath);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("创建目标目录失败：" + ex.Message);
                    }
                }
                //获得源文件下所有文件
                List<string> files = new List<string>(Directory.GetFiles(sourcePath));
                files.ForEach(c =>
                {
                    string destFile = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                    //覆盖模式
                    if (File.Exists(destFile))
                    {
                        File.Delete(destFile);
                    }
                    if(progressBar.Value <= 90)
                    {
                        progressBar.Value += 10;
                    }
                    File.Copy(c, destFile);
                });
                //获得源文件下所有目录文件
                List<string> folders = new List<string>(Directory.GetDirectories(sourcePath));

                folders.ForEach(c =>
                {
                    string destDir = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                    MoveFolder(c, destDir, progressBar);
                });
            }
            else
            {
                throw new DirectoryNotFoundException("源目录不存在！");
            }
        }
        /// <summary>
        /// 创建桌面快捷方式
        /// </summary>
        /// <param name="FileName">文件的名称</param>
        /// <param name="exePath">EXE的路径</param>
        /// <returns>成功或失败</returns>
        public static bool CreateDesktopShortcut(string FileName, string exePath)
        {
            try
            {
                string deskTop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
                if (File.Exists(deskTop + FileName + ".lnk"))  //
                {
                    File.Delete(deskTop + FileName + ".lnk");//删除原来的桌面快捷键方式
                }
                WshShell shell = new WshShell();

                //快捷键方式创建的位置、名称
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(deskTop + FileName + ".lnk");
                shortcut.TargetPath = exePath; //目标文件                         
                DirectoryInfo pathInfo = new DirectoryInfo(exePath);
                string WorkingDirectory = pathInfo.Parent.FullName;  //该属性指定应用程序的工作目录，当用户没有指定一个具体的目录时，快捷方式的目标应用程序将使用该属性所指定的目录来装载或保存文件。
                shortcut.WorkingDirectory = WorkingDirectory;
                shortcut.WindowStyle = 1; 
                shortcut.Description = FileName; //描述
                shortcut.Save(); //必须调用保存快捷才成创建成功
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
    public class PrinterOperate
    {
        public static void GetPrinters (Label label, ComboBox combox)
        {
            PrintDocument print = new PrintDocument();
            string sDefault = print.PrinterSettings.PrinterName;//默认打印机名
            label.Text = "默认打印机:" + sDefault;
            combox.Text = sDefault;
            foreach (string sPrint in PrinterSettings.InstalledPrinters)//获取所有打印机名称
            {
                combox.Items.Add(sPrint);
            }
        }

        [DllImport("winspool.drv")]
        public static extern bool SetDefaultPrinter(String Name); //调用win api将指定名称的打印机设置为默认打印机
    }
    public class NetworkOperate
    {
        public static string GetLocalIP()
        {
            IPAddress localIp = null;
            try
            {
                IPAddress[] ipArray;
                ipArray = Dns.GetHostAddresses(Dns.GetHostName());
                localIp = ipArray.First(ip => ip.AddressFamily == AddressFamily.InterNetwork);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace + "\r\n" + ex.Message, "错误", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
            if (localIp == null)
            {
                localIp = IPAddress.Parse("127.0.0.1");
            }
            return localIp.ToString();
        }
        /// <summary>
        /// Ping命令检测网络是否畅通
        /// </summary>
        /// <param name="urls">URL数据</param>
        /// <param name="errorCount">ping时连接失败个数</param>
        /// <returns></returns>
        public static bool Ping(string url)
        {
            bool isconn = true;
            Ping ping = new Ping();
            try
            {
                PingReply pr;
                pr = ping.Send(url);
                if (pr.Status != IPStatus.Success)
                {
                    isconn = false;
                }
            }
            catch
            {
                isconn = false;
            }
            //if (errorCount > 0 && errorCount < 3)
            //  isconn = true;
            return isconn;
        }
    }
}
