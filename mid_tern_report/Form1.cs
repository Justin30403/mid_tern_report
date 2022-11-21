using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace mid_tern_report
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // 打開程式後，先跳出Form2，確認帳號密碼後 才進入主程式(Form1)
            Form2 form2;
            form2 = new Form2();
            form2.ShowDialog();
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

            richTextBox1.Text = String.Format("{0}物品為:{1} , 單價:{2} * 數量:{3}={4}元 ", _radiobutton_log, _combobox_log,_price,_num, _price * _num);


            DataGridViewRowCollection rows = dataGridView1.Rows;
            DateTime date = DateTime.Now; // 現在時間
            rows.Add(new Object[] {"",date.ToString("yyyy/MM/dd HH:mm:ss"),_radiobutton_log,_combobox_log,_price,_num,_price*_num });

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
           
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
