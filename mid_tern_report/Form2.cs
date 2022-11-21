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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.ControlBox = false; // 直接隱藏form2的關閉視窗按鈕
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        public bool check = false;
        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("登入");
            String username = "test";
            String password = "1234";
            /*
            if (textBox1.Text.Equals(username) && textBox2.Text.Equals(password))
            {
                check = true;
                this.Close();
            }
            else
            {
                MessageBox.Show("帳號或密碼錯了");
            }
            
            */
            if (String.Equals(username, textBox1.Text))
            {
                if (String.Equals(password, textBox2.Text))
                {
                    this.Close();
                }
                else
                {
                    MessageBox.Show("密碼錯誤");
                }
            }
            else
            {
                MessageBox.Show("帳號錯誤");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 關閉form2
            // this.Close();
            // 關閉整個application
            System.Environment.Exit(0);

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
