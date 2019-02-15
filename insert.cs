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
    public partial class insert : Form
    {
        public insert()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == string.Empty)
            {
                MessageBox.Show("请输入产品名称"); return;
            }
            if (textBox1.Text == string.Empty)
            {
                MessageBox.Show("请输入单位名称");return;
            }
            if (textBox3.Text == string.Empty)
                textBox3.Text = "0";
            if (textBox4.Text == string.Empty)
                textBox4.Text = "0";
            if (comboBox1.Text == string.Empty)
                comboBox1.Text = "未分类";
            if (pd())
            {
                int id ;
                string idsql = @"select count(*) from [" + Form1.id + "]";
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
                conn.Open();
                OleDbCommand idcmd = new OleDbCommand(idsql, conn);
                id = Convert.ToInt32(idcmd.ExecuteScalar())+1;
                string sql = @"insert into [" + Form1.id + "](id,bianhao,product,danwei,price1,price2,beizhu) values("+id+",'"+textBox5.Text+"','" + textBox2.Text + "','"+textBox1.Text+"','" + textBox3.Text + "','" + textBox4.Text + "','" + comboBox1.Text + "')";                
                OleDbCommand cmd = new OleDbCommand(sql, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("添加成功");
                this.Hide();
                guanli.way = 1;
                guanli.mark = 1;
            }
            else
                MessageBox.Show("该分类已存在相同产品");
        }
        public bool pd()//检测product和beizhu是否会出现相同值
        {
            string sql = @"select * from [" + Form1.id + "] where [product]='" + textBox2.Text + "' and [beizhu]='" + comboBox1.Text + "'";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
                return false;
            else
                return true;
        }
        public void showfenlei()
        {
            string sql = @"select [beizhu] from [" + Form1.id + "] group by [beizhu]";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            comboBox1.DataSource = ds;
            comboBox1.DisplayMember = "" + Form1.id + ".beizhu";
            conn.Close();
        }

        private void insert_Load(object sender, EventArgs e)
        {
            showfenlei();
        }
    }
}
