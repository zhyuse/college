using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
namespace 财务销售
{
    public partial class Form1 : Form
    {
        public static string id;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(jiance())
            {
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");//C: \Users\Administrator\Desktop\财务\Database1.mdb
                conn.Open();
                //select *from 表名 where 字段名='字段值';*表示全表，从全表中
                //OleDbCommand cmd = new OleDbCommand("select * from [user]", conn);  //表名要用[]括起来
                //label1.Text = cmd.ExecuteScalar().ToString();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * from [user] where id=" + textBox1.Text + " and password='" + textBox2.Text + "'", conn);
                //DataSet ds = new DataSet();
                //da.Fill(ds);
                //MessageBox.Show(ds.Tables[0].Rows[0]["name"].ToString());
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    id = textBox1.Text;
                    mainpage form = new mainpage();
                    form.Show();
                    this.Hide();
                }
                else
                    MessageBox.Show("密码错误或用户名不存在");
                conn.Close();
            }           
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            register form = new register();
            form.Show();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)//keychar 是ascii码,enter对应为13
                button1.PerformClick();
        }
        public bool jiance()
        {
            if (textBox1.Text == string.Empty || textBox2.Text == string.Empty )
            {
                MessageBox.Show("检测到有没输入的选项，请输入完整");
                return false;
            }
            return true;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)//keychar 是ascii码,enter对应为13
                button1.PerformClick();
        }

    }
}
