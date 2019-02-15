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
    public partial class register : Form
    {
        public register()
        {
            InitializeComponent();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void register_Load(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if(jiance())
            {
                string sql1 = @"create table [" + textBox1.Text + "] ([product] VARCHAR(45) NOT NULL, [price1] int null,[price2] int null,[beizhu] varchar(45) null ); ";
                string sql2=@"insert into [user]([id],[name],[password]) values('"+textBox1.Text+"','"+textBox2.Text+"','"+textBox3.Text+"')";
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
                conn.Open();
                OleDbCommand cmd1 = new OleDbCommand(sql1, conn);
                OleDbCommand cmd2 = new OleDbCommand(sql2, conn);
                cmd1.ExecuteNonQuery();
                cmd2.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("注册成功");
                this.Close();
            }
        }
        public bool jiance()
        {
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            if (textBox1.Text == string.Empty || textBox2.Text == string.Empty || textBox3.Text == string.Empty || textBox4.Text == string.Empty)
            {
                MessageBox.Show("检测到有没输入的选项，请输入完整");
                return false;
            }
            if (textBox3.Text != textBox4.Text)
            {
                MessageBox.Show("密码不一致");
                return false;
            }
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * from [user] where id=" + textBox1.Text + " ", conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("该用户名已经被注册了");
                return false;
            }
            return true;
        }
    }
}
