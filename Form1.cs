using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Excel;

//using ExcelDataReader;

using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;


using System.Collections.Specialized;

namespace ty5_2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();


            // 设置窗体和textBox_Link允许拖放
            this.AllowDrop = true;
            this.textBox1_Link.AllowDrop = true;
            this.textBox2_Target.AllowDrop = true;


            // 为两个文本框添加KeyDown事件处理
            this.textBox1_Link.KeyDown += new KeyEventHandler(TextBox1_KeyDown);
            this.textBox2_Target.KeyDown += new KeyEventHandler(TextBox2_KeyDown);

            // 为textBox_Link和textBox_Target添加事件处理程序
            this.textBox1_Link.DragEnter += new DragEventHandler(textBox_DragEnter);
            this.textBox1_Link.DragDrop += new DragEventHandler(textBox_DragDrop);
            this.textBox2_Target.DragEnter += new DragEventHandler(textBox_DragEnter);
            this.textBox2_Target.DragDrop += new DragEventHandler(textBox_DragDrop);


            this.linkLabel1.LinkClicked += new LinkLabelLinkClickedEventHandler(LinkLabel1_LinkClicked);


            this.backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            this.backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);


            //// 重定向 Console 的输出到 textBox3_Log
            TextWriter writer = new TextBoxWriter(textBox3_Log);
            Console.SetOut(writer);
            Console.SetError(writer); // 重定向错误输出


        }


        private void TextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            // 检测 Ctrl+V 组合键
            if (e.Control && e.KeyCode == Keys.V)
            {
                // 检查剪贴板中是否有文件路径
                if (Clipboard.ContainsFileDropList())
                {
                    // 获取剪贴板中的文件路径
                    StringCollection files = Clipboard.GetFileDropList();
                    if (files.Count > 0)
                    {
                        string firstPath = files[0]; // 获取第一个路径

                        // 检查文件是否为 .xls 文件
                        if (File.Exists(firstPath) &&
                            (Path.GetExtension(firstPath).Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                             Path.GetExtension(firstPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)))
                        {
                            // 如果是 .xls 或 .xlsx 文件，接受它
                            textBox1_Link.Text = firstPath;
                        }
                        else
                        {
                            // 如果不是 .xls 文件，显示提示
                            // 异步显示消息提示，避免锁死文件浏览窗口
                            this.BeginInvoke(new MethodInvoker(delegate
                            {
                                MessageBox.Show("请选择.xls文件");
                            }));
                            return; // 终止处理
                        }

                        e.Handled = true; // 阻止默认的粘贴行为
                    }
                }
            }
        }


        private void TextBox2_KeyDown(object sender, KeyEventArgs e)
        {
            // 检测 Ctrl+V 组合键
            if (e.Control && e.KeyCode == Keys.V)
            {
                // 使用完全限定的类型名称解决命名冲突
                System.Windows.Forms.TextBox textBox = sender as System.Windows.Forms.TextBox;
                if (textBox != null)
                {
                    // 检查剪贴板中是否有文件路径
                    if (Clipboard.ContainsFileDropList())
                    {
                        // 获取剪贴板中的文件路径
                        StringCollection files = Clipboard.GetFileDropList();
                        if (files.Count > 0)
                        {
                            string firstPath = files[0]; // 获取第一个路径

                            // 判断是文件还是目录
                            if (File.Exists(firstPath))
                            {
                                // 如果是文件，获取其所在目录
                                string directory = Path.GetDirectoryName(firstPath);
                                textBox.Text = directory;
                            }
                            else if (Directory.Exists(firstPath))
                            {
                                // 如果是目录，直接使用该路径
                                textBox.Text = firstPath;
                            }

                            e.Handled = true; // 阻止默认的粘贴行为
                        }
                    }
                }
            }
        }


        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 将链接标记为已访问
            this.linkLabel1.LinkVisited = true;

            // 使用浏览器打开链接
            System.Diagnostics.Process.Start(this.linkLabel1.Text);
        }

        // 创建一个安全更新控件的TextWriter类
        public class TextBoxWriter : TextWriter
        {
            private TextBox _textBox;
            private StringBuilder _buffer;

            public TextBoxWriter(TextBox textBox)
            {
                _textBox = textBox;
                _buffer = new StringBuilder();
            }

            public override void Write(char value)
            {
                _buffer.Append(value);
                if (value == '\n')
                {
                    Flush();
                }
            }

            public override void Flush()
            {
                _textBox.BeginInvoke(new Action(() =>
                {
                    _textBox.AppendText(_buffer.ToString());
                    _buffer.Remove(0, _buffer.Length);
                }));
            }

            public override Encoding Encoding => Encoding.UTF8;
        }




        private void textBox_DragEnter(object sender, DragEventArgs e)
        {
            // 检查拖动数据是否包含文件
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy; // 允许复制操作
            }
            else
            {
                e.Effect = DragDropEffects.None; // 不允许其他类型的拖动操作
            }
        }

        private void textBox_DragDrop(object sender, DragEventArgs e)
        {
            // 获取拖动的文件数组
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            // 如果是文件，根据文件类型更新文本框内容
            if (files != null && files.Length > 0)
            {
                // 由于sender是object类型，我们需要将其转换为TextBox
                TextBox textBox = (TextBox)sender;

                if (sender == textBox1_Link)
                {
                    // 检查文件是否为.xls扩展名
                    if (!Path.GetExtension(files[0]).ToLower().Equals(".xls") && !Path.GetExtension(files[0]).ToLower().Equals(".xlsx"))
                    {
                        // 异步显示消息提示，避免锁死文件浏览窗口
                        this.BeginInvoke(new MethodInvoker(delegate
                        {
                            MessageBox.Show("请选择.xls文件");
                        }));
                        return; // 终止处理
                    }

                    textBox.Text = files[0]; // 添加文件路径


                    openFileDialog1.FileName = files[0];

                }
                else if (sender == textBox2_Target)
                {
                    // 获取文件所在目录
                    string directory = Path.GetDirectoryName(files[0]);
                    textBox.Text = directory; // 设置为文件所在目录

                    folderBrowserDialog1.SelectedPath = directory;
                }
            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            SetLinkLabel1TextAsync("http://sasmap.mmc199.com");
            CheckFileAndFolder();
        }

        private void SetLinkLabel1TextAsync(string url)
        {
            WebClient client = new WebClient();
            client.DownloadStringCompleted += (sender, e) =>
            {
                if (e.Error == null)
                {
                    linkLabel1.Text = e.Result;
                }
                else
                {
                    linkLabel1.Text = "https://github.com/mmc199/sasmap/releases/latest/download/ty5-2.xls";
                }
                client.Dispose();
            };
            client.DownloadStringAsync(new Uri(url));
        }


        private void CheckFileAndFolder()
        {
            string appPath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(appPath, "ty5-2.xls");
            string folderPath = Path.Combine(appPath, "Maps\\ty5-2");

            // 检查文件是否存在并更新UI
            if (File.Exists(filePath))
            {
                textBox1_Link.Text = filePath;
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(filePath);
                openFileDialog1.FileName = Path.GetFileName(filePath);
            }

            // 如果文件夹不存在则创建
            Directory.CreateDirectory(folderPath);

            // 更新UI
            textBox2_Target.Text = folderPath;
            folderBrowserDialog1.SelectedPath = folderPath;
        }










        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "图源xls|*.xls;*.xlsx";

            // 检查 textBox1_Link 中的路径是否存在
            if (File.Exists(textBox1_Link.Text))
            {
                // 设置 OpenFileDialog 的初始目录为 textBox1_Link 中的路径
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(textBox1_Link.Text);
            }
            else
            {
                // 如果 textBox1_Link 中的路径不存在，可以设置为默认路径或给出提示
                openFileDialog1.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
                // MessageBox.Show("指定的路径不存在，将使用默认目录。");
            }

            // 显示 OpenFileDialog
            DialogResult result = openFileDialog1.ShowDialog();

            // 如果用户选择了文件，更新 textBox1_Link 的内容
            if (result == DialogResult.OK && !string.IsNullOrEmpty(openFileDialog1.FileName))
            {
                textBox1_Link.Text = openFileDialog1.FileName;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            // 检查 textBox2_Target 中的路径是否存在
            if (Directory.Exists(textBox2_Target.Text))
            {
                // 设置 FolderBrowserDialog 的初始目录为 textBox2_Target 中的路径
                folderBrowserDialog1.SelectedPath = textBox2_Target.Text;
            }

            // 显示 FolderBrowserDialog
            DialogResult result = folderBrowserDialog1.ShowDialog();

            // 如果用户选择了文件夹，更新 textBox2_Target 的内容
            if (result == DialogResult.OK && !string.IsNullOrEmpty(folderBrowserDialog1.SelectedPath))
            {
                textBox2_Target.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {


            bool error = false;

            if (textBox1_Link.Text == "" || textBox2_Target.Text == "") // 检查是否在未做任何操作的情况下点击了提交
            {
                error = true;
                MessageBox.Show("必须指定文件/目录。");
            }
            if (!Directory.Exists(textBox2_Target.Text)) // 检查目标目录是否存在
            {
                    try
                    {
                        // 创建目录及其所有父目录
                        Directory.CreateDirectory(textBox2_Target.Text);

                        // 可选：显示成功创建的消息
                        MessageBox.Show($"已成功创建目标目录: {textBox2_Target.Text}");
                    }
                    catch (Exception ex)
                    {
                        // 如果创建失败，显示错误消息
                        error = true;
                        MessageBox.Show($"无法创建目标目录: {textBox2_Target.Text}\n错误信息: {ex.Message}");
                    }
                
            }
            if (!File.Exists(textBox1_Link.Text)) // 检查xls文件是否存在
            {
                error = true;
                MessageBox.Show("您在“xls”文本框中指定的文件不存在。");
            }
            // 如果没有错误，我们现在创建实际链接
            if (error == false)
            {

                //string scriptDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
                //Console.WriteLine($"Script directory: {scriptDir}");

                // 获得Excel文件路径
                //string excelPath = Path.Combine(scriptDir, "ty5-2.xls");
                string excelPath = textBox1_Link.Text;
                string scriptDir = textBox2_Target.Text;

                Console.WriteLine($"Excel path: {excelPath}");

                // 打开excel文件，并创建一个stream
                using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
                {
                    // 使用ExcelDataReader创建一个 reader
                    // 需要根据文件扩展名决定使用哪种reader
                    IExcelDataReader reader;
                    if (Path.GetExtension(excelPath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (Path.GetExtension(excelPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        throw new Exception("Unsupported file format");
                    }

                    using (reader)
                    {
                        // 读取excel文件中的所有数据到一个数据集dataset中
                        var result = reader.AsDataSet();

                        // 从数据集中取第一个表
                        DataTable dt = result.Tables[0];
                        Console.WriteLine($"Table columns: {dt.Columns.Count}, Table rows: {dt.Rows.Count}");

                        // 创建一个字典用于存放数据
                        var dataDict = new Dictionary<string, List<string>>();

                        // 从第二行开始遍历该表格的每一行，因为一般第一行是表头
                        for (int i = 3; i < dt.Rows.Count; i++)
                        {
                            Console.WriteLine($"Processing row: {i}");

                            // 列遍历
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                // 如果当前列的值不为空，并且该列的第一个元素包含'='字符
                                if (dt.Rows[i][j] != DBNull.Value && dt.Rows[2][j].ToString().Contains('=') && !dt.Rows[i][1].ToString().Contains('N'))
                                {
                                    Console.WriteLine($"Condition met for row {i}, column {j}");
                                    // 获取该行的第0列的值，作为键
                                    string key = dt.Rows[i][0].ToString().Trim();
                                    // 获取当前格的值，作为值
                                    string value = dt.Rows[i][j].ToString();

                                    // 如果字典中不含有key，则添加
                                    if (!dataDict.ContainsKey(key))
                                    {
                                        Console.WriteLine($"Adding key to dictionary: {key}");
                                        dataDict.Add(key, new List<string>());
                                    }
                                    // 将value添加到字典相应的key上
                                    Console.WriteLine($"Adding value to key {key}: {value}");
                                    dataDict[key].Add(value);
                                }
                            }
                        }

                        Console.WriteLine($"Data Dictionary count: {dataDict.Count}");

                        // 在文件系统中遍历这个字典
                        foreach (var item in dataDict)
                        {
                            Console.WriteLine($"Processing params dictionary item: Key - {item.Key}, Values count - {item.Value.Count}");

                            // 为每一个键创建一个同名的目录
                            string directoryName = $"{item.Key}.zmp";
                            string directoryPath = Path.Combine(scriptDir, directoryName);
                            Console.WriteLine($"Directory path: {directoryPath}");

                            // 如果该目录不存在，则创建
                            if (!Directory.Exists(directoryPath))
                            {
                                Console.WriteLine($"Creating directory: {directoryPath}");
                                Directory.CreateDirectory(directoryPath);
                            }

                            // 在新创建的目录下创建一个params.txt文件，并写入对应的value列表，之间用换行分隔
                            string filePath = Path.Combine(directoryPath, "params.txt");
                            Console.WriteLine($"File path: {filePath}");

                            string[] lines = item.Value.ToArray();

                            File.WriteAllLines(filePath, lines, System.Text.Encoding.Default);
                            Console.WriteLine($"Wrote lines to file: {filePath}");
                            File.AppendAllText(filePath, "\n\n", System.Text.Encoding.Default);
                            Console.WriteLine($"Appended text to file: {filePath}");
                        }


                        // 从数据集中取第二个表
                        DataTable dt2 = result.Tables[1];
                        Console.WriteLine($"Table2 columns: {dt2.Columns.Count}, Table2 rows: {dt2.Rows.Count}");

                        // 创建一个字典用于存放第二个表的数据
                        var getUrlDict = new Dictionary<string, List<string>>();

                        // 从第二行开始遍历该表格的每一行，因为一般第一行是表头
                        for (int i = 3; i < dt2.Rows.Count; i++)
                        {
                            Console.WriteLine($"Processing row: {i}");
                            // 列遍历
                            for (int j = 0; j < dt2.Columns.Count; j++)
                            {
                                // 如果当前列的值不为空，并且该列的第一个元素包含'='字符
                                if (dt2.Rows[i][j] != DBNull.Value && !dt2.Rows[2][j].ToString().Contains('=') && !dt.Rows[i][1].ToString().Contains('N'))
                                {
                                    Console.WriteLine($"Condition met for row {i}, column {j}");

                                    // 获取该行的第0列，也就是'URL'列的值，作为键
                                    string key = dt2.Rows[i][0].ToString().Trim();
                                    // 获取当前格的值，作为值
                                    string value = dt2.Rows[i][j].ToString();

                                    // 如果字典中不含有key，则添加
                                    if (!getUrlDict.ContainsKey(key))
                                    {
                                        Console.WriteLine($"Adding key to Url dictionary: {key}");
                                        getUrlDict.Add(key, new List<string>());
                                    }
                                    // 将value添加到字典相应的key上
                                    Console.WriteLine($"Adding value to Url key {key}: {value}");
                                    getUrlDict[key].Add(value);
                                }
                            }
                        }

                        Console.WriteLine($"GetUrl Dictionary count: {getUrlDict.Count}");

                        // 在文件系统中遍历这个getUrl字典
                        foreach (var item in getUrlDict)
                        {
                            Console.WriteLine($"Processing getUrl dictionary item: Key - {item.Key}, Values count - {item.Value.Count}");

                            // 为每一个键创建一个同名的目录
                            string getUrlDirectoryName = $"{item.Key}.zmp";
                            string getUrlDirectoryPath = Path.Combine(scriptDir, getUrlDirectoryName);
                            Console.WriteLine($"Url Directory path: {getUrlDirectoryPath}");

                            // 如果该目录不存在，则创建
                            if (!Directory.Exists(getUrlDirectoryPath))
                            {
                                Console.WriteLine($"Creating Url directory: {getUrlDirectoryPath}");
                                Directory.CreateDirectory(getUrlDirectoryPath);
                            }

                            // 在新创建的目录下创建一个GetUrlScript.txt文件，并写入对应的value列表，之间用换行分隔
                            string getUrlFilePath = Path.Combine(getUrlDirectoryPath, "GetUrlScript.txt");
                            Console.WriteLine($"Url file path: {getUrlFilePath}");

                            string[] getUrlLines = item.Value.ToArray();

                            File.WriteAllLines(getUrlFilePath, getUrlLines, System.Text.Encoding.Default);
                            Console.WriteLine($"Wrote lines to Url file: {getUrlFilePath}");
                            File.AppendAllText(getUrlFilePath, "\n\n", System.Text.Encoding.Default);
                            Console.WriteLine($"Appended text to Url file: {getUrlFilePath}");
                        }


                    }


                }


            }


        }

        private void button4_Click(object sender, EventArgs e)
        {
            string downloadUrl = linkLabel1.Text;
            string fileName = "ty5-2.xls";
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            if (File.Exists(filePath))
            {
                if (MessageBox.Show("本地已经存在 " + fileName + " 文件, 是否要覆盖?", "确认覆盖", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                        Path.GetFileNameWithoutExtension(fileName) + DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetExtension(fileName));
                }
            }

            button4.Text = "下载中…";
            //UpdateButtonText("下载中…");
            button4.Enabled = false; // 禁用下载按钮

            backgroundWorker1.RunWorkerAsync(new string[] { downloadUrl, filePath });
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] args = e.Argument as string[];
            string downloadUrl = args[0];
            string filePath = args[1];

            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                ServicePointManager.ServerCertificateValidationCallback = ValidateServerCertificate;

                string redirectedUrl = GetRedirectedUrl(downloadUrl);


                DownloadFile(redirectedUrl, filePath);


                e.Result = new DownloadResult { Success = true, FilePath = filePath };
            }
            catch (Exception ex)
            {
                e.Result = new DownloadResult { Success = false, ErrorMessage = ex.Message };
            }
        }


        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            button4.Text = "下载模板";
            //UpdateButtonText("下载模板");
            button4.Enabled = true; // 重新启用下载按钮

            DownloadResult result = e.Result as DownloadResult;
            if (result.Success)
            {
                MessageBox.Show("文件下载成功。");
                textBox1_Link.Text = result.FilePath;
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(result.FilePath);
                openFileDialog1.FileName = Path.GetFileName(result.FilePath);
            }
            else
            {
                MessageBox.Show("下载过程中出现异常: " + result.ErrorMessage);
            }
        }

        private string GetRedirectedUrl(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.AllowAutoRedirect = false;
            request.Method = "GET";

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                return (response.StatusCode == HttpStatusCode.MovedPermanently ||
                    response.StatusCode == HttpStatusCode.Found ||
                    response.StatusCode == HttpStatusCode.Redirect) ?
                    response.Headers["Location"] : url;
            }
        }

        private void DownloadFile(string url, string outputPath)
        {
            using (var client = new WebClient())
            {
                client.DownloadFile(url, outputPath);
            }
        }

        private static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        private class DownloadResult
        {
            public bool Success { get; set; }
            public string FilePath { get; set; }
            public string ErrorMessage { get; set; }
        }

        //private void UpdateButtonText(string text)
        //{
        //    if (button4.InvokeRequired)
        //    {
        //        button4.Invoke(new Action<string>(UpdateButtonText), text);

        //        MessageBox.Show("Invoke成功。");
        //    }
        //    else
        //    {
        //        button4.Text = text;
        //    }
        //}

    }

}
