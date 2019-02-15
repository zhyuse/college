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
    public partial class dayin2 : Form
    {
        public dayin2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double price;
            if (!pd())
            { MessageBox.Show("销售单中已存在相同的产品");return; }
            if (textBox1.Text == string.Empty)
            { MessageBox.Show("请输入产品名称");return; }
            if (textBox2.Text == string.Empty)
            { MessageBox.Show("请输入单位名称");return; }
            if (textBox3.Text == string.Empty)
            { MessageBox.Show("请输入数量"); return; }
            if (textBox2.Text == string.Empty)
            { MessageBox.Show("请输入单价"); return; }
            price = Convert.ToDouble(textBox3.Text) * Convert.ToDouble(textBox4.Text);
            int id;
            string idsql = @"select count(*) from [dayin] ";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            OleDbCommand idcmd = new OleDbCommand(idsql, conn);
            id = Convert.ToInt32(idcmd.ExecuteScalar()) + 1;
            string sql = @"insert into [dayin](id,product,danwei,[number],danjia,price) values(" + id + ",'" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "'," + price+ ")";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            dayin.fanye = 1;//判断dayin的datagridview1是否可以翻下一页
            this.Hide();
            dayin.mark = 1;
        }
        public bool pd()//检测product和beizhu是否会出现相同值
        {
            string sql = @"select * from [dayin] where [product]='" +textBox1.Text + "' ";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
                return false;
            else
                return true;
        }

        private void dayin2_Load(object sender, EventArgs e)
        {
            textBox1.Text = dayin.product;
            textBox2.Text = dayin.danwei;
            textBox3.Text = "1";
            textBox4.Text = dayin.price2.ToString();
        }
    }
}
