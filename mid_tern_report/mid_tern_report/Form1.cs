using Microsoft.Office.Interop.Excel;
using ScottPlot;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ScottPlot.Plottable;




namespace mid_tern_report
{
    public partial class Form1 : Form
    {
        // ScottPlot start
        private Crosshair Crosshair;
        // ScottPlot end

        


            public Form1()
        {
            InitializeComponent();

            // 打開程式後，先跳出Form2，確認帳號密碼後 才進入主程式(Form1)
            Form2 form2;
            form2 = new Form2();
            form2.ShowDialog();
            // ScottPlot start
            Crosshair = formsPlot1.Plot.AddCrosshair(0, 0);
            formsPlot1.Refresh();
            // ScottPlot end



        }

        private void button1_Click(object sender, EventArgs e)
        {


            string i_input_price = textBox1.Text;
            string i_input_num = textBox2.Text;
            double _price = Convert.ToDouble(i_input_price);
            double _num = Convert.ToDouble(i_input_num);


            string _radiobutton_log = "";
            if (radioButton2.Checked == true)  //如果選出貨 顯示出貨 如果不選的話預設為進貨
            { _radiobutton_log = "出貨"; }
            else
            { _radiobutton_log = "進貨"; }

            string _combobox_log = comboBox1.SelectedItem.ToString();
            
            richTextBox1.Text = String.Format("{0}物品為:{1} , 單價:{2} * 數量:{3}={4}元 ", _radiobutton_log, _combobox_log, _price, _num, _price * _num);
          

             DataGridViewRowCollection rows = dataGridView1.Rows;
                DateTime date = DateTime.Now; // 現在時間
                rows.Add(new Object[] {" ",date.ToString("yyyy/MM/dd HH:mm:ss"), _radiobutton_log, _combobox_log, _price, _num, _price * _num }); ; ;




            if (radioButton1.Checked == true)
            {
                string y = textBox2.Text;
                string x = comboBox1.Text;
                this.chart3.Series["product"].Points.AddXY(x, y);
                this.chart4.Series["product"].Points.AddXY(x, y);
            }
            else 
            {
                string y = textBox2.Text;
                string x = comboBox1.Text;
                this.chart1.Series["product"].Points.AddXY(x, y);
                this.chart2.Series["product"].Points.AddXY(x, y);
            }

            
            
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
         
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is about的資訊");
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            
            DataGridViewCellCollection selRowData = dataGridView1.Rows[e.RowIndex].Cells;

            string _type = "";
            _type = Convert.ToString(selRowData[2].Value);

            if (_type.Equals("進貨"))
            {
                radioButton1.Checked = true;
            }
            else
            {
                radioButton2.Checked = true;
            }


            this.comboBox1.Text = Convert.ToString(selRowData[3].Value);
            this.textBox1.Text = Convert.ToString(selRowData[4].Value);
            this.textBox2.Text = Convert.ToString(selRowData[5].Value);
            this.label5.Text = Convert.ToString(selRowData[0].Value);

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
        
        private void button3_Click_1(object sender, EventArgs e)
        {
            
           
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart1_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart1.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart_Data";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件

            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把資料塞進 Excel 內

                // 標題
                Sheet.Cells[1, 1] = "標籤";
                Sheet.Cells[1, 2] = "數量";

                // 內容
                for (int k = 0; k < this.chart1.Series["stocks"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart1.Series["stocks"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart1.Series["stocks"].Points[k].YValues[0].ToString();
                }

                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {
            String[] n = { "第一期", "第二期", "第三期", "第四期", "全年" };
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart_Data";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件

            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把資料塞進 Excel 內

                // 標題
                Sheet.Cells[1, 1] = "標籤";
                Sheet.Cells[1, 2] = "數量";

                // 內容
                for (int k = 0; k < this.chart1.Series["stocks"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart1.Series["stocks"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart1.Series["stocks"].Points[k].YValues[0].ToString();
                }

                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart1_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart1.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void chart4_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
         
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
             double[] values = ScottPlot.DataGen.RandomWalk(100);
            //double[] values = [1.1, 2.2,...,9.8];

            // 先看一下第五個元素是什麼
            MessageBox.Show(values[5].ToString());

            // 2. 用plt這個變數，當作【圖表數據】的捷徑
            var plt = formsPlot1.Plot;
            //var plt = new ScottPlot.Plot(600, 400);

            double[] xs = DataGen.Consecutive(51);
            double[] sin = DataGen.Sin(51);
            double[] cos = DataGen.Cos(51);

           plt.AddScatter(xs, sin, color: Color.Red);
            plt.Clear();
            plt.AddScatter(xs, cos, color: Color.Blue);

            plt.SaveFig("quickstart_clear.png");
            ////////////////////////////////
            // 3. 開始繪圖
            //formsPlot1_MouseLeave(null, null);
           // plt.AddSignal(values);

            // Set axis limits to control the view
            // (min x, max x, min y, max y)
            //plt.SetAxisLimits(0, 100, -25, 25);

            // 3. 繪圖結束
            ////////////////////////////////

            // 4. 將統計圖顯示在GUI上面
            formsPlot1.Refresh();


        }
        private void formsPlot1_MouseMove(object sender, MouseEventArgs e)
        {
            (double coordinateX, double coordinateY) =
                                                 formsPlot1.GetMouseCoordinates();

            Crosshair.X = coordinateX;
            Crosshair.Y = coordinateY;

            formsPlot1.Refresh(lowQuality: true, skipIfCurrentlyRendering: true);
        }

        // 滑鼠移動進入圖表時，顯示座標
        private void formsPlot1_MouseEnter(object sender, EventArgs e)
        {
            Crosshair.IsVisible = true;
        }

        // 滑鼠移動離開圖表時，關閉顯示座標
        private void formsPlot1_MouseLeave(object sender, EventArgs e)
        {
            Crosshair.IsVisible = false;
            formsPlot1.Refresh();
        }


        private void button13_Click(object sender, EventArgs e)
        {
           

        }

        private void formsPlot1_Load(object sender, EventArgs e)
        {

        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_bar_Chart_Data";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件

            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把資料塞進 Excel 內

                // 標題
                Sheet.Cells[1, 1] = "文具";
                Sheet.Cells[1, 2] = "數量";
                
                // 內容
                for (int k = 0; k < this.chart1.Series["product"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart1.Series["product"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart1.Series["product"].Points[k].YValues[0].ToString();

                }


                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_bar_Chart1_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart1.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);

        }

        private void button10_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_bar_Chart_Data";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件

            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把資料塞進 Excel 內

                // 標題
                Sheet.Cells[1, 1] = "文具";
                Sheet.Cells[1, 2] = "占比";

                // 內容
                for (int k = 0; k < this.chart1.Series["product"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart1.Series["product"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart1.Series["product"].Points[k].YValues[0].ToString();
                }


                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }
        }
    }
}

