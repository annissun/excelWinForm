using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace excelDo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folder = new System.Windows.Forms.FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                this.label1.Text = folder.SelectedPath;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folder = new System.Windows.Forms.FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                this.label2.Text = folder.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //获取文件夹下面所有子目录
            String pathIn = this.label1.Text;
            String pathOut = this.label2.Text;
            getAllFiles(pathIn, pathOut);
        }
        public void getAllFiles(String pathin, String pathOut)
        {
            if (pathOut == "" || pathOut.IndexOf("输出路径")>=0) {
                pathOut = pathin;
            }
            //第一层，获取地区
            DirectoryInfo[] citys = getFoders(pathin);
            foreach (DirectoryInfo city in citys)
            {
                //displayLog("地区名称为:"+city.Name);
                //第二层，获取驾校名称
                DirectoryInfo[] carNames = getFoders(city.FullName);
                foreach (DirectoryInfo carName in carNames)
                {
                    //displayLog("驾校名称为:"+carName.Name);
                    //第三层，获取日期
                    DirectoryInfo[] dates = getFoders(carName.FullName);
                    foreach (DirectoryInfo date in dates)
                    {
                        //displayLog("日期名称为:" + date.Name);
                        //日期文件夹下面所有的文件
                        FileInfo[] fileNames = getFilesName(date.FullName);
                        foreach (FileInfo fileName in fileNames)
                        {
                            String[] sArray = date.Name.Split('.');
                            String datenow = "2019-"+sArray[0] + "-" + sArray[1];
                            displayLog("文件名称为:" + fileName.Name + "  日期为:"+ date.Name+"  驾校名称为:"+ carName.Name+"  地区名称为:"+ city.Name);
                            List<OneData> myData = ExcelToDatat(fileName.FullName);
                            foreach (OneData data in myData) {
                                data.carName = carName.Name;
                                data.date = datenow;
                            }
                            displayLog("读取完成"+ fileName.FullName);
                            //将文件写入新的xls表
                            //WriteExcel(myData, pathOut);
                        }
                    }
                }
            }

        }

        //将数据插入表格
        private void WriteExcel(List<OneData> myData, String pathOut) {
            if (myData.Count <= 0) {
                return;
            }
            String fileName = "output.xls";
            DirectoryInfo root = new DirectoryInfo(pathOut);
            FileInfo[] filesNames = root.GetFiles(fileName);
            if (filesNames.Length <= 0)
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  
                using (FileStream file = new FileStream(pathOut + "\\" + fileName, FileMode.Create))
                {
                    workbook.Write(file);  //创建test.xls文件。
                    file.Close();
                }
                WriteExcel(myData, pathOut);
            }
            else {
                DataToExcel(myData, pathOut+ "\\" + fileName);
            }

        }

        //excel转化为数据
        private List<OneData> ExcelToDatat(string filePath)
        {

            IWorkbook wk = null;
            List<OneData> DataAll = new List<OneData>();
            string extension = System.IO.Path.GetExtension(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    wk = new XSSFWorkbook(fs);
                }
                fs.Close();
                //读取当前表数据
                ISheet sheet = wk.GetSheetAt(0);

                IRow row = sheet.GetRow(0);  //读取当前行数据
                                             //LastRowNum 是当前表的总行数-1（注意）
                int nameCol = -1; //第几列是姓名 
                int idCol = -1;  //第几列是身份证
                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    row = sheet.GetRow(i);  //读取当前行数据
                    if (row != null)
                    {
                        OneData data = new OneData();
                        //LastCellNum 是当前行的总列数
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            //读取该行的第j列数据，数据为空则为null
                            ICell hssfCell = row.GetCell(j);
                            if (hssfCell == null)
                                continue;
                            String name = hssfCell.StringCellValue;
                            if (name.IndexOf("保险人") >= 0 || name.IndexOf("姓名") >= 0 || name.IndexOf("名称") >= 0 || name.IndexOf("名字") >= 0)
                            {
                                nameCol = j;
                            }
                            else if (name.IndexOf("证") >= 0 && name.IndexOf("号码") >= 0)
                            {
                                idCol = j;
                            }
                            else
                            {
                                if (j == nameCol)
                                {
                                    data.name = hssfCell.StringCellValue;
                                }
                                else if (j == idCol)
                                {
                                    data.idCard = hssfCell.StringCellValue;
                                }
                            }
                        }
                        if (data.name != null && data.idCard != null && data.name != "" && data.idCard != "") {
                            DataAll.Add(data);
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                //只在Debug模式下才输出
                //richTextBox1.AppendText(ex.Message + "\n");
            }
            return DataAll;
        }
        
        //数据转化为Excel
        private static void DataToExcel(List<OneData> myData, String pathOut)
        {
            FileStream fs = File.OpenRead(pathOut);
            IWorkbook wk = new HSSFWorkbook(fs);
            fs.Close();
            ISheet sheet = wk.GetSheetAt(0);
            IRow row = null;
            ICell cell = null;
            int allRow = sheet.LastRowNum+1;
            foreach (OneData data in myData)
            {
                row = sheet.CreateRow(allRow);
                for (int j = 0; j < 4; j++)
                {
                    cell = row.CreateCell(j);//excel第二行开始写入数据
                    if (j == 0) {
                       cell.SetCellValue(data.name);
                    }
                    else if (j == 1) {
                        cell.SetCellValue(data.idCard);
                    }
                    else if (j == 2)
                    {
                        cell.SetCellValue(data.carName);
                    }
                    else if (j == 3)
                    {
                        cell.SetCellValue(data.date);
                    }
                }
                allRow += 1;
            }

            using (FileStream fileStream = File.Open(pathOut,
            FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fileStream);
                fileStream.Close();
            }
        }

        //输出日志到屏幕
        public void displayLog(String log)
        {
            this.textBox1.AppendText(log + "\r\n");
        }

        public DirectoryInfo[] getFoders(String path) {
            DirectoryInfo root = new DirectoryInfo(path);
            DirectoryInfo[] list = root.GetDirectories();
            return list;
        }
        /*
         * 获得指定路径下所有文件名
         * StreamWriter sw  文件写入流
         * string path      文件路径
         * int indent       输出时的缩进量
         */
        public FileInfo[] getFilesName(string path)
        {
            DirectoryInfo root = new DirectoryInfo(path);
            FileInfo[] filesName = root.GetFiles("*.xls");
            return filesName;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
