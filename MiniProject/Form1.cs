using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;
using System.Collections; 

namespace MiniProject
{

    

/*=====================================================================*/
    public partial class Form1 : Form
    {
        public double factor=0.0086; 
        ReadDataFile reading=new ReadDataFile();
        private int clickCount = 0;
        string FileName; 
       
        public Form1()
        {
            InitializeComponent();
          //  this.textBox1.Text= "Difference";
            this.Text = "DATA ANALYSIS THE DIFFERENCE BETWEEN PEAKS ";

        }
/*==============================THIS IS PROCESS BUTTON =======================*/
     private void button1_Click(object sender, EventArgs e)
        {
            
                   
               clickCount++; 
               reading.Readtext3(FileName);
               /*=======================================================*/
                string f = this.textBox4.Text.ToString();

                double factor =Convert.ToDouble(this.numericUpDown1.Value.ToString()); 
                int p1 = reading.setMinPeak();
                int p2 = reading.setMaxPeak();
                this.textBox2.Text = p1.ToString();// peak1.ToString();
                this.textBox3.Text = p2.ToString();
                this.textBox4.Text = factor.ToString(); 

                double Difference = reading.Peaksdifference(factor);
                this.textBox1.Text = Difference.ToString();

               
                    reading.WriteInExcel(Difference, clickCount);    //Column 
                    Console.WriteLine("This file {0}", clickCount);
                    reading.Bank0Data.Clear(); // New added 
               
             //******************************************************* 
              DrawChart(reading.BankAll.ToArray());
             
            
        }

  
 
   private void DrawChart(int[] DataIn)
            {

                chart1.Series.Clear();
                Refresh();

                this.chart1.ChartAreas[0].AxisY.Crossing = 0;

                chart1.ChartAreas[0].AxisX.Title = "Values";
                chart1.ChartAreas[0].AxisX.Maximum = DataIn.Length;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                chart1.ChartAreas[0].AxisX.Interval = 42;
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                if (chart1.Series.IsUniqueName("Data"))
                {
                    chart1.Series.Add("Data");
                }
              
                chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                for (int i = 0; i <= DataIn.Length - 1; i++)
                    this.chart1.Series["Data"].Points.AddXY(i, DataIn[i]);



            }
   /*==============================THIS IS CLOSE BUTTON =======================*/
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close(); 
        }

        private void Form1_Load(object sender, EventArgs e)
        {
          
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
             Random rnd = new Random();
            string nChartImage = "chart1" + rnd.Next(1, 13).ToString()+".png";

            this.chart1.SaveImage(nChartImage, ChartImageFormat.Png); 
        }
  /*==============================THIS IS LOAD BUTTON =======================*/
        private void button4_Click(object sender, EventArgs e)
        {
         string Pathdata; 
         System.Windows.Forms.OpenFileDialog openFileDialog1;
         openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
         openFileDialog1 = new OpenFileDialog();
         openFileDialog1.ShowDialog();
         openFileDialog1.Title = "Load  Data File";
         openFileDialog1.DefaultExt = "txt";
         Pathdata = openFileDialog1.FileName;
         FileName = Path.GetFileName(Pathdata);
         Console.WriteLine("the file name is ={0}", FileName); 

        }
    /*==============================THIS IS RESTART BUTTON =======================*/
        private void button5_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            Refresh();
            reading.Bank0Data.Clear();
            reading.BankAll.Clear();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            clickCount = 0;  

        }

    }
    /*==========================================================================================
     * ReadDataFile Class reads that data from Text files , Save data in Excel Sheet 1 and Sheet2
     * Calculates the minimum and maximum data  peaks and the difference between them.
     * 
     * *****************************************************************************************/
    class ReadDataFile
    {
        public List<int> Bank0Data = new List<int>();
        public List<int> BankAll = new List<int>();
        int peak1, peak2;
        int addSheet = 0;
        
  /*======================================================*/
        public int setMaxPeak()
        {
            peak1 = -1; 
            int maxIndex = -1;
            double maxInt = Int32.MinValue;
             
            
            for (int i = 0; i < BankAll.Count; i++)
            {
                int value = BankAll[i];
                if (value > maxInt)
                {
                    maxInt = value;
                    maxIndex = i;
                }
            }
            peak1 = maxIndex;
            return peak1;
        }
        /*======================================================*/
        public int setMinPeak()
        {
            peak2 = -1; 
            int minIndex = -1;
            double minInt = Convert.ToDouble(Int32.MaxValue);
            
            
            for (int i = 0; i < BankAll.Count; i++)
            {
                int value = BankAll[i];
                if (value < minInt)
                {
                    minInt = value;
                    minIndex = i;
                }
            }
            peak2 = minIndex;
            return peak2;
        }
        /*======================================================*/
        public double Peaksdifference(double factor)
        {
            return ((double)peak2 - (double)peak1) * factor;
        }

        /*======================================================*/
        public void Readtext3(string fileName)
        {
            string fileContent1 = File.ReadAllText(fileName);
            string[] integerStrings1 = fileContent1.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int Maxlength = integerStrings1.Length;
            int val;
            for (int n = 0; n < Maxlength; n++)
            {
                val = int.Parse(integerStrings1[n]);
                Bank0Data.Add(val);
                BankAll.Add(val); 
                Console.WriteLine("{0} \t {1} ", n, Bank0Data[n]);
            }

            Console.WriteLine("Number of elements (rows) is ={0}", Bank0Data.Count);
            Console.WriteLine(" Data type of {0} ", Bank0Data.GetType());

          

        }



        public void WriteInExcel(double Difference, int column)
        {

            string filename = "File" + column.ToString();
            Console.WriteLine("1- The size of this array={0}", Bank0Data.Count.ToString());
            var spreadsheetLocation = Path.Combine(Directory.GetCurrentDirectory(), "sdata_test.xlsx");
            var exApplication = new Microsoft.Office.Interop.Excel.Application();
            var exWorkbook = exApplication.Workbooks.Open(spreadsheetLocation);
           // var   exWorksheet = exWorkbook.Sheets.Add(After: exWorkbook.Sheets[1]);

            var exWorksheet = (Excel.Worksheet)exWorkbook.Worksheets.get_Item(1);
            exWorksheet.Cells[column, 1] = Difference;
            if (addSheet == 0)
            {
                exWorksheet = exWorkbook.Sheets.Add(After: exWorkbook.Sheets[1]);
                addSheet++; 
            }
            else exWorksheet = (Excel.Worksheet)exWorkbook.Worksheets.get_Item(2);

            for (var row = 2; row <= Bank0Data.Count - 1; row++)
            {
                var cell = (Excel.Range)exWorksheet.Cells[row, column];
                cell.Value2 = Bank0Data[row];
            }
            var header = (Excel.Range)exWorksheet.Cells[1, column];
            header.Value2 = "File"+column;

            Console.WriteLine("2- The size of this array={0}", Bank0Data.Count.ToString());
            exWorkbook.Save();
            GC.Collect();
            GC.WaitForPendingFinalizers();
           //  Marshal.ReleaseComObject(writeRange);
            Marshal.ReleaseComObject(exWorksheet);
            //close and release and  
            //quit and release

            exApplication.Quit();
            Marshal.ReleaseComObject(exApplication);
            //Console.WriteLine("3- The size of this array={0}", Bank0Data.Count.ToString());

        }


    }   // END OF THE CLASS    // 
    /*======================================================================*/
}
