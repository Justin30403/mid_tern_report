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
using System.Data.SQLite;
using ScottPlot;
using ScottPlot.Plottable;

using iText.Forms;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Extgstate;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Pdf.Canvas.Draw;
using MathNet.Numerics.Statistics;



namespace mid_tern_report
{
    public partial class Form1 : Form
    {

        //int index = 1;

        public class DBConfig
        {
            //log.db要放在【bin\Debug底下】      
            public static string dbFile = Application.StartupPath + @"\log.db";

            public static string dbPath = "Data source=" + dbFile;

            public static SQLiteConnection sqlite_connect;
            public static SQLiteCommand sqlite_cmd;
            public static SQLiteDataReader sqlite_datareader;
        }

        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        private void Show_DB()
        {
            this.dataGridView1.Rows.Clear();

            string sql = @"SELECT * from record;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    int _serial = Convert.ToInt32(DBConfig.sqlite_datareader["serial"]);
                    int _date = Convert.ToInt32(DBConfig.sqlite_datareader["date"]);
                    int _type = Convert.ToInt32(DBConfig.sqlite_datareader["type"]);
                    string _name = Convert.ToString(DBConfig.sqlite_datareader["name"]);
                    double _price = Convert.ToDouble(DBConfig.sqlite_datareader["price"]);
                    double _number = Convert.ToDouble(DBConfig.sqlite_datareader["number"]);
                    double _total = _price * _number;

                    string _date_str = DateTimeOffset.FromUnixTimeSeconds(_date).ToString("yy-MM-dd hh:mm:ss");

                    string _type_str = "";
                    if (_type == 0)
                    { _type_str = "進貨"; }
                    else { _type_str = "出貨"; }

                    i = _serial;
                    DataGridViewRowCollection rows = dataGridView1.Rows;
                    rows.Add(new Object[] { i, _date_str, _type_str, _name, _price, _number
                                               , _total });
                }
                DBConfig.sqlite_datareader.Close();
            }
        }
        private Crosshair Crosshair;

        int i = 1;


        public Form1()
        {
            InitializeComponent();

            // 打開程式後，先跳出Form2，確認帳號密碼後 才進入主程式(Form1)
            Form2 form2;
            form2 = new Form2();
            form2.ShowDialog();
            //test
            // ScottPlot start
            Crosshair = formsPlot1.Plot.AddCrosshair(0, 0);
            formsPlot1.Refresh();
            // ScottPlot end


           
            this.label5.Text = i.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
           


            string i_input_price = textBox1.Text;
            string i_input_number = textBox2.Text;
            double _price = Convert.ToDouble(i_input_price);
            double _number = Convert.ToDouble(i_input_number);

            string _radiobutton_log = "";
            if (radioButton2.Checked == true)  //如果選出貨 顯示出貨 如果不選的話預設為進貨
            { _radiobutton_log = "出貨"; }
            else
            { _radiobutton_log = "進貨"; }

            string _combobox_log = comboBox1.SelectedItem.ToString();

            richTextBox1.Text = String.Format("{0}物品為:{1} , 單價:{2} * 數量:{3}={4}元 ", _radiobutton_log, _combobox_log, _price, _number, _price * _number);


            DataGridViewRowCollection rows = dataGridView1.Rows;
            DateTime date = DateTime.Now; // 現在時間
            rows.Add(new Object[] { i, date.ToString("yyyy/MM/dd HH:mm:ss"), _radiobutton_log, _combobox_log, _price, _number, _price * _number });



            if (radioButton1.Checked == true)
            {
                string y = textBox2.Text;
                string x = comboBox1.Text;
                this.chart2.Series["PRODUCT"].Points.AddXY(x, y);
                this.chart4.Series["PRODUCT"].Points.AddXY(x, y);
            }
            else
            {
                string y = textBox2.Text;
                string x = comboBox1.Text;
                this.chart1.Series["PRODUCT"].Points.AddXY(x, y);
                this.chart3.Series["PRODUCT"].Points.AddXY(x, y);
            }

            i += 1;
            this.label5.Text = i.ToString();

            

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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            PrintPDF();

        }
        void PrintPDF()
        {
            // Set the output dir and file name
            // string directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            string src = "./test.pdf";
            string dst = @"./new_test.pdf";

            //manipulatePdf(src, dst);

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); ;
            saveFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                dst = saveFileDialog.FileName;
                manipulatePdf(src, dst);
            }
        }
        void manipulatePdf(String src, String dst)
        {

            // read image
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); ;
            openFileDialog.Filter = "All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            string image_file = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                image_file = openFileDialog.FileName;
            }
            PdfWriter writer = new PdfWriter(dst);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            PdfFont font_tr = PdfFontFactory.CreateFont(@"c:/Windows/fonts/kaiu.ttf", PdfEncodings.IDENTITY_H);
            PdfPage page = pdf.AddNewPage();
            PdfPage page1 = pdf.GetPage(1);
            PdfCanvas pdfCanvas1 = new PdfCanvas(page1);
            iText.Kernel.Geom.Rectangle rectangle = new iText.Kernel.Geom.Rectangle(200, 700, 100, 100);
            iText.Layout.Canvas canvas = new iText.Layout.Canvas(pdfCanvas1, rectangle);
            ImageData imageData = ImageDataFactory.Create(image_file);
            iText.Layout.Element.Image image = new iText.Layout.Element.Image(imageData);
            canvas.Add(image);
            ImageData imageData2 = ImageDataFactory.Create(image_file);
            iText.Layout.Element.Image image2 = new iText.Layout.Element.Image(imageData2);
            image2.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            image2.SetHeight(60);
            image2.SetWidth(120);
            image2.SetMarginLeft(130);
            image2.SetMarginTop(400);
            document.Add(image2);
            canvas.Close();
            document.Close();

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
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

        private void textBox1_TextChanged(object sender, EventArgs e)
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

        private void button3_Click(object sender, EventArgs e)
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
                for (int k = 0; k < this.chart1.Series["PRODUCT"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart1.Series["PRODUCT"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart1.Series["PRODUCT"].Points[k].YValues[0].ToString();
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

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart1_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart3.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void button7_Click(object sender, EventArgs e)
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
                Sheet.Cells[1, 3] = "百分比";

                // 內容
                for (int k = 0; k < this.chart3.Series["PRODUCT"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart3.Series["PRODUCT"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart3.Series["PRODUCT"].Points[k].YValues[0].ToString();
                    Sheet.Cells[k + 2, 3] = this.chart3.Series["PRODUCT"].Points[k].YValues.ToString();

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

        private void button5_Click(object sender, EventArgs e)
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
                for (int k = 0; k < this.chart2.Series["PRODUCT"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart2.Series["PRODUCT"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart2.Series["PRODUCT"].Points[k].YValues[0].ToString();
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
                Sheet.Cells[1, 3] = "百分比";

                // 內容
                for (int k = 0; k < this.chart4.Series["PRODUCT"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart4.Series["PRODUCT"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart4.Series["PRODUCT"].Points[k].YValues[0].ToString();
                    Sheet.Cells[k + 2, 3] = this.chart4.Series["PRODUCT"].Points[k].YValues[0].ToString();
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

        private void button6_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart1_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart2.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);

        }

        private void button10_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart1_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart4.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            System.Drawing.Bitmap bitmap = null;
            //let string to qr-code
            string strQrCodeContent = richTextBox1.Text;

            // QR Code產生器
            ZXing.BarcodeWriter writer = new ZXing.BarcodeWriter
            {
                Format = ZXing.BarcodeFormat.QR_CODE,
                Options = new ZXing.QrCode.QrCodeEncodingOptions
                {
                    //Create Photo
                    Height = 200,
                    Width = 200,
                    CharacterSet = "UTF-8",

                    //錯誤修正容量
                    //L水平    7%的字碼可被修正
                    //M水平    15%的字碼可被修正
                    //Q水平    25%的字碼可被修正
                    //H水平    30%的字碼可被修正
                    ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H
                }

            };
            //Create Qr-code , use input string
            bitmap = writer.Write(strQrCodeContent);
            //Storage bitmpa

            string strDir;
            strDir = Directory.GetCurrentDirectory();
            strDir += "\\temp.jpg";
            bitmap.Save(strDir, System.Drawing.Imaging.ImageFormat.Jpeg);
            //Display to picturebox
            pictureBox1.Image = bitmap;


        }


        private void button14_Click(object sender, EventArgs e)
        {


        }

        private void chart4_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {

            // 1. 先產生一維陣列，共有1000000個數字
            double[] values = ScottPlot.DataGen.RandomWalk(1000);
            //double[] values = [1.1, 2.2,...,9.8];

            // 先看一下第五個元素是什麼
            MessageBox.Show(values[5].ToString());

            // 2. 用plt這個變數，當作【圖表數據】的捷徑
            var plt = formsPlot1.Plot;

            ////////////////////////////////
            // 3. 開始繪圖

            // sample rate: 一個刻度有幾個點
            //  e.g., 現在共有1000個點，每個刻度200個點，所以有五個刻度
            //var plt = new ScottPlot.Plot(600, 400);

            formsPlot1_MouseLeave(null, null);
            plt.AddSignal(values);

            // Set axis limits to control the view
            // (min x, max x, min y, max y)
            plt.SetAxisLimits(0, 100, -25, 25);

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

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            System.Drawing.Bitmap bitmap = null;
            //let string to qr-code
            string strQrCodeContent = richTextBox1.Text;

            // QR Code產生器
            ZXing.BarcodeWriter writer = new ZXing.BarcodeWriter
            {
                Format = ZXing.BarcodeFormat.QR_CODE,
                Options = new ZXing.QrCode.QrCodeEncodingOptions
                {
                    //Create Photo
                    Height = 200,
                    Width = 200,
                    CharacterSet = "UTF-8",

                    //錯誤修正容量
                    //L水平    7%的字碼可被修正
                    //M水平    15%的字碼可被修正
                    //Q水平    25%的字碼可被修正
                    //H水平    30%的字碼可被修正
                    ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H
                }

            };
            //Create Qr-code , use input string
            bitmap = writer.Write(strQrCodeContent);
            //Storage bitmpa

            string strDir;
            strDir = Directory.GetCurrentDirectory();
            strDir += "\\temp.jpg";
            bitmap.Save(strDir, System.Drawing.Imaging.ImageFormat.Jpeg);
            //Display to picturebox
            pictureBox1.Image = bitmap;

        }

        private void button14_Click_1(object sender, EventArgs e)
        {


        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }


        private void button14_Click_4(object sender, EventArgs e)
        {
            List<double> data = new List<double>();
            data.Add(3);
            data.Add(6);
            data.Add(12);
            data.Add(35);
            data.Add(2);
            data.Add(14);
            data.Add(66);
            data.Add(44);
            data.Add(23);
            data.Add(69);

            // 2. get statistic data
            string _log = statistic(data);

            // 3. get qr code
            System.Drawing.Bitmap bitmap = get_qrcode(_log);

            // 4. export to pdf
            export_to_pdf(bitmap);

        }
        void export_to_pdf(Bitmap bitmap)
        {
            // Set the output dir and file name

            string dst = @"./new_test.pdf";


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); ;
            saveFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                dst = saveFileDialog.FileName;

            }


            // 1. create pdf
            PdfWriter writer = new PdfWriter(dst);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            // 標楷體
            PdfFont font_tr = PdfFontFactory.CreateFont(@"c:/Windows/fonts/kaiu.ttf", PdfEncodings.IDENTITY_H);

            // add picture
            PdfPage page = pdf.AddNewPage();
            PdfPage page1 = pdf.GetPage(1);
            PdfCanvas pdfCanvas1 = new PdfCanvas(page1);
            iText.Kernel.Geom.Rectangle rectangle = new iText.Kernel.Geom.Rectangle(200, 700, 100, 100);
            iText.Layout.Canvas canvas = new iText.Layout.Canvas(pdfCanvas1, rectangle);
            ImageData imageData = ImageDataFactory.Create(BmpToBytes(bitmap));
            iText.Layout.Element.Image image = new iText.Layout.Element.Image(imageData);
            canvas.Add(image);

            // 4. close content
            canvas.Close();
            document.Close();

        }

        //Bitmap to Byte array
        public byte[] BmpToBytes(Bitmap bmp)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            byte[] b = ms.GetBuffer();
            return b;
        }

        public Bitmap get_qrcode(string log)
        {
            System.Drawing.Bitmap bitmap = null;
            //let string to qr-code
            string strQrCodeContent = log;

            ZXing.BarcodeWriter writer = new ZXing.BarcodeWriter
            {
                Format = ZXing.BarcodeFormat.QR_CODE,
                Options = new ZXing.QrCode.QrCodeEncodingOptions
                {
                    //Create Photo
                    Height = 200,
                    Width = 200,
                    CharacterSet = "UTF-8",

                    //錯誤修正容量
                    //L水平    7%的字碼可被修正
                    //M水平    15%的字碼可被修正
                    //Q水平    25%的字碼可被修正
                    //H水平    30%的字碼可被修正
                    ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H
                }

            };
            //Create Qr-code , use input string
            bitmap = writer.Write(strQrCodeContent);
            /*
            string strDir;
            strDir = Directory.GetCurrentDirectory();
            strDir += "\\temp.jpg";
            bitmap.Save(strDir, System.Drawing.Imaging.ImageFormat.Jpeg);
            */
            return bitmap;
        }

        // 輸出統計數據
        public string statistic(List<double> data)
        {
            double mean = Statistics.Mean(data);
            double stddiv = Statistics.StandardDeviation(data);
            double pstddiv = Statistics.PopulationStandardDeviation(data);
            double variance = Statistics.Variance(data);
            double median = Statistics.Median(data);
            double lowerQuartile = Statistics.LowerQuartile(data);
            double upperQuartile = Statistics.UpperQuartile(data);
            double interQuartileRange = Statistics.InterquartileRange(data);
            double min = Statistics.Minimum(data);
            double max = Statistics.Maximum(data);

            string _log = "";
            _log = string.Format("平均值: {0}\n " +
                "標準差: {1}\n" +
                "變異數: {2}\n" +
                "中位數: {3}\n" +
                "最小值: {4}\n" +
                "最大值: {5}",
                mean, stddiv, variance, median, min, max);

            return _log;

        }

        private void button2_Click_1(object sender, EventArgs e)
        {


        }

        private void button11_Click_1(object sender, EventArgs e)
        {
         
        }

    }
}

