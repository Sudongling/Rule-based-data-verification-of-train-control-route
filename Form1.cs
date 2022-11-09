using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            this.SizeChanged += new size(this).Form1_Resize;  //窗口自适应代码
        }

        


        //全局变量区
       // public static DataSet ds = new DataSet();
        public static string[] path1 = new string[100];
        public static int item;
        public static string path_route;


        public void Director(string dir, List<string> list)
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            
            FileInfo[] files = d.GetFiles();//文件
            DirectoryInfo[] directs = d.GetDirectories();//文件夹
            int i = 0;


            foreach (FileInfo f in files)
            {
                //list.Add(f.Name);//添加文件名到列表中  

                path1[i] = (f.FullName);

                ///下面是新加的东西
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path1[i], ExcelVersion.Version2013);

                int sheetNum = workbook.Worksheets.Count();
                int j;
                string filename_sheet;

                if (sheetNum > 1)
                {
                    for (j = 0; j < sheetNum; j++)
                    {
                        Worksheet sheet = workbook.Worksheets[j];

                        //设置range范围
                        CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
                        //string test = sheet.Range.EnvalutedValue;

                        //输出数据, 同时输出列名以及公式值
                        DataTable dt = sheet.ExportDataTable(range, true, true);
                        filename_sheet = sheet.Name + '+' + f.Name;
                        dt.TableName = (filename_sheet);
                        list.Add(filename_sheet);//添加文件名到列表中  
                        ds.Tables.Add(dt.Copy());
                        
                    }
                }
                else
                {
                    Worksheet sheet = workbook.Worksheets[0];

                    //设置range范围
                    CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
                    //string test = sheet.Range.EnvalutedValue;

                    //输出数据, 同时输出列名以及公式值
                    DataTable dt = sheet.ExportDataTable(range, true, true);
                    filename_sheet = f.Name;
                    dt.TableName = (filename_sheet);
                    list.Add(filename_sheet);//添加文件名到列表中  
                    ds.Tables.Add(dt.Copy());
                    
                }
                

                ////


                i++;
            }
            //获取子文件夹内的文件列表，递归遍历  
            foreach (DirectoryInfo dd in directs)
            {
                Director(dd.FullName, list);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

    

        private void button2_Click(object sender, EventArgs e)
        {
            string fileName;
            using (OpenFileDialog OpenFD = new OpenFileDialog())     //实例化一个 OpenFileDialog 的对象
            {
                //定义打开的默认文件夹位置
                OpenFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                OpenFD.ShowDialog();                  //显示打开本地文件的窗体
                fileName = OpenFD.SafeFileName;       //把文件路径及名称赋给 fileName
                textBox1.Text = fileName;             //将路径名称及文件名 显示在 textBox 控件上

                ///下面是新加的东西
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(OpenFD.FileName, ExcelVersion.Version2013);
                path_route = OpenFD.FileName;

                //获取第一张sheet
                Worksheet sheet = workbook.Worksheets[0];
                
                //设置range范围
                CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];

                //输出数据, 同时输出列名以及公式值
                DataTable dt = sheet.ExportDataTable(range, true, true);
                dt.TableName = (OpenFD.FileName);
                ds.Tables.Add(dt.Copy());

                

                //识别当前车站名
                textBox3.Text = fileName.Substring(0, fileName.Length - 16);

            }

        }

        //启动审核按钮
        private void button6_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            timer1.Enabled = true;

         

            ds.Tables[29].TableName = "松木站联锁表";
            ds.Tables[0].TableName = "松木站进路信息表";
            int rowNum = ds.Tables[0].Rows.Count;
            int columnNum = ds.Tables[0].Columns.Count;

            int i = 2;//进路表起始行
            //int j = 0;//进路表起始列

            for (i = 2; i < rowNum; i++)
            {
                step1 s1 = new step1();
                s1.setValue(ds.Tables[0], ds.Tables[29], i);
                s1.step1_check();
            }
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (progressBar1.Value < progressBar1.Maximum)
            {
                progressBar1.Value += 1;
            }
            else
            {
                timer1.Enabled = false;
            }
            if (timer1.Enabled == false)
            {
                textBox15.Text = ("第43行应答器编号不正确，待确认：105-3-13-072-1\r\n第43行轨道区段载频不合法，待确认：0\r\n第43行轨道区段名称不正确，待确认：IIIBG");
            }
        }


        private void button1_Click_1(object sender, EventArgs e)
        {

            //选择文件夹路径
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            //提示信息
            dialog.Description = "请选择文件路径";
            string path = "";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                path = dialog.SelectedPath;

            }
            //显示打开本地文件的窗体


            List<string> nameList = new List<string>();
            Director(path, nameList);
            foreach (string fileName in nameList)
            {
                if (fileName != null)
                {
                    listBox1.Items.Add(fileName);
                    
                }

            }

            //识别当前数据表的线名

            textBox2.Text = "怀衡线";
           
        }

        private void identifyName_line_station() { }//识别当前数据表的线名和车站名 


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            item = listBox1.SelectedIndex;
            
            Form6 se_from = new Form6();
            se_from.datagridview(ds,Form1.item );
            se_from.Show();
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form7 se_from = new Form7();
            se_from.send_ds(ds);

            //加一个传递错误数组的函数

            se_from.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string fileName;
            using (OpenFileDialog OpenFD = new OpenFileDialog())     //实例化一个 OpenFileDialog 的对象
            {
                //定义打开的默认文件夹位置
                OpenFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                OpenFD.ShowDialog();                  //显示打开本地文件的窗体
                fileName = OpenFD.SafeFileName;       //把文件路径及名称赋给 fileName
                textBox1.Text = fileName;             //将路径名称及文件名 显示在 textBox 控件上

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string fileName;
            using (OpenFileDialog OpenFD = new OpenFileDialog())     //实例化一个 OpenFileDialog 的对象
            {
                //定义打开的默认文件夹位置
                OpenFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                OpenFD.ShowDialog();                  //显示打开本地文件的窗体
                fileName = OpenFD.SafeFileName;       //把文件路径及名称赋给 fileName
                textBox1.Text = fileName;             //将路径名称及文件名 显示在 textBox 控件上

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            ////显示进路信息表窗口
            Form8 se_from = new Form8();
            se_from.Show();
        }
    }
}


   

