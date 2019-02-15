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
    public partial class guanli : Form
    {
        public static double price1, price2;
        public static string bianhao,danwei, product,beizhu;
        public static int mark=0,way=1;//way-选择查询方式
        public static int m = 0, page, nowpage=1,allnum;//m-已查询的行数
        public guanli()
        {
            InitializeComponent();
        }

        private void guanli_Load(object sender, EventArgs e)
        {
            //datagridview 选定整行,属性selectmode-fullrowselect
            dataGridView1.AllowUserToResizeColumns = false;//锁定列宽
            dataGridView1.AllowUserToResizeRows = false;//锁定行高
            showall();
            label3.Text = "总共有" + page + "页，当前第" + nowpage + "页";
            fenlei();
        }
        public void showall()
        {
            dataGridView1.AllowUserToAddRows = false;//删掉最后一行
            dataGridView1.RowHeadersVisible = false;//删掉左边一列
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string sqlnum= @"select count(*) from ["+Form1.id+"]";
            OleDbCommand cmdnum = new OleDbCommand(sqlnum, conn);
            allnum = Convert.ToInt32(cmdnum.ExecuteScalar());
            if (allnum % 17 == 0)
                page = allnum / 17;
            else
                page = allnum / 17 + 1;
            string sql;
            //OleDbDataAdapter da = new OleDbDataAdapter("SELECT * from ["+Form1.id+"] ", conn);
            if (m == 0)
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] ORDER BY [beizhu], [product]";
            else
                sql= "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] where id not in(select top "+m+" id from ["+Form1.id+ "] order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds,""+Form1.id+"");
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        public void showfenlei()
        {
            dataGridView1.AllowUserToAddRows = false;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string sqlnum = @"select count(*) from [" + Form1.id + "] where [beizhu]='" + comboBox1.Text + "' ";
            OleDbCommand cmdnum = new OleDbCommand(sqlnum, conn);
            allnum = Convert.ToInt32(cmdnum.ExecuteScalar());
            if (allnum % 17 == 0)
                page = allnum / 17;
            else
                page = allnum / 17 + 1;
            string sql;
            if (m == 0)
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] where [beizhu]='" + comboBox1.Text + "' ORDER BY [beizhu], [product]";
            else
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] where [beizhu]='"+comboBox1.Text+"' and id not in(select top " + m + " id from [" + Form1.id + "] where [beizhu]='" + comboBox1.Text + "' order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        public void showabout()
        {
            dataGridView1.AllowUserToAddRows = false;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string text = "%" + textBox1.Text + "%";//暂时用这种方法来实现模糊查询
            string sqlnum;
            sqlnum = @"select count(*) from [" + Form1.id + "] where [product] like '" + text + "' ";
            OleDbCommand cmdnum = new OleDbCommand(sqlnum, conn);
            allnum = Convert.ToInt32(cmdnum.ExecuteScalar());
            if (allnum % 17 == 0)
                page = allnum / 17;
            else
                page = allnum / 17 + 1;
            string sql;
            if(m!=0)
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [product] like '"+text+ "' and id not in(select top " + m + " id from [" + Form1.id + "] where [product] like '" + text + "' order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            else
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [product] like '" + text + "' ORDER BY [beizhu], [product]";
            //access模糊查询要用*
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        public void showbianhao()
        {
            dataGridView1.AllowUserToAddRows = false;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string text = "%" + textBox2.Text + "%";//暂时用这种方法来实现模糊查询
            string sqlnum;
            sqlnum = @"select count(*) from [" + Form1.id + "] where [bianhao] like '" + text + "' ";
            OleDbCommand cmdnum = new OleDbCommand(sqlnum, conn);
            allnum = Convert.ToInt32(cmdnum.ExecuteScalar());
            if (allnum % 17 == 0)
                page = allnum / 17;
            else
                page = allnum / 17 + 1;
            string sql;
            if (m != 0)
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [bianhao] like '" + text + "' and id not in(select top " + m + " id from [" + Form1.id + "] where [bianhao] like '" + text + "' order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            else
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price1] as 进货价,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [bianhao] like '" + text + "' ORDER BY [beizhu], [product]";
            //access模糊查询要用*
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            if (dataGridView1.CurrentRow != null)
            {//return;
                DataGridViewRow dgvr = dataGridView1.CurrentRow;
                bianhao= dgvr.Cells["编号"].Value.ToString();
                product = dgvr.Cells["产品"].Value.ToString();
                danwei = dgvr.Cells["单位"].Value.ToString();
                price1 = Convert.ToDouble(dgvr.Cells["进货价"].Value);
                price2 = Convert.ToDouble(dgvr.Cells["售价"].Value);
                beizhu = dgvr.Cells["分类"].Value.ToString();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(mark==1)
            {
                showall();
                fenlei();
                mark--;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)//增加
        {
            insert form = new insert();
            form.Show();
        }

        private void button2_Click(object sender, EventArgs e)//修改
        {
            update form = new update();
            form.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            nowpage = 1;
            m = 0;
            way = 3;
            showbianhao();
            上一页.Enabled = false;
            if (page > 1)
                下一页.Enabled = true;
            else
                下一页.Enabled = false;
            label3.Text = "总共有" + page + "页，当前第" + nowpage + "页";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)//实现实时修改datagridview的text传值
        {
            if (dataGridView1.CurrentRow != null)
            {//return;
                DataGridViewRow dgvr = dataGridView1.CurrentRow;
                bianhao = dgvr.Cells["编号"].Value.ToString();
                product = dgvr.Cells["产品"].Value.ToString();
                danwei = dgvr.Cells["单位"].Value.ToString();
                price1 = Convert.ToDouble(dgvr.Cells["进货价"].Value);
                price2 = Convert.ToDouble(dgvr.Cells["售价"].Value);
                beizhu = dgvr.Cells["分类"].Value.ToString();
            }
        }

        private void 上一页_Click(object sender, EventArgs e)
        {
            m -= 17;
            label6.Text = m.ToString();
            if (way == 1)
                showall();
            else if (way == 2)
                showfenlei();
            else if (way == 3)
                showabout();
            if (m == 0)
                上一页.Enabled = false;
            if (m < allnum - 17)
                下一页.Enabled = true;
            nowpage--;
            label3.Text = "总共有" + page + "页，当前第" + nowpage + "页";
        }

        private void 下一页_Click(object sender, EventArgs e)
        {
            m += 17;
            label6.Text = m.ToString();
            if (way == 2)
                showfenlei();
            else if (way == 1)
                showall();
            else if (way == 3)
                showabout();
            if (m >= allnum - 17)
                下一页.Enabled = false;
            if (m > 0)
                上一页.Enabled = true;
            nowpage++;
            label3.Text= "总共有" + page + "页，当前第" + nowpage + "页";
        }

        private void button4_Click(object sender, EventArgs e)//分类查询
        {
            nowpage = 1;
            m = 0;
            if (comboBox1.Text != string.Empty)
            { showfenlei(); way = 2; }
            else
            { showall(); way = 1; }
            上一页.Enabled = false;
            if (page > 1)
                下一页.Enabled = true;
            else
                下一页.Enabled = false;
            label3.Text = "总共有" + page + "页，当前第" + nowpage + "页";
        }

        private void button5_Click(object sender, EventArgs e)//模糊查询
        {
            nowpage = 1;
            m = 0;
            way = 3;
            showabout();
            上一页.Enabled = false;
            if (page > 1)
                下一页.Enabled = true;
            else
                下一页.Enabled = false;
            label3.Text = "总共有" + page + "页，当前第" + nowpage + "页";
        }

        private void button3_Click(object sender, EventArgs e)//删除
        {
            if (DialogResult.Cancel != MessageBox.Show("确认删除？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information))
            {
                int id;
                string idsql = @"select id from [" + Form1.id + "] where product='" + product + "' and beizhu='" + beizhu + "'";
                string sql = @"delete from [" + Form1.id + "] where bianhao='"+bianhao+"' and product='" + product + "' and beizhu='" + beizhu + "'";
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
                conn.Open();
                OleDbCommand idcmd = new OleDbCommand(idsql, conn);
                id = Convert.ToInt32(idcmd.ExecuteScalar());
                OleDbCommand cmd = new OleDbCommand(sql, conn);
                cmd.ExecuteNonQuery();
                string updateidsql = @"update [" + Form1.id + "] set [id]=[id]-1 where [id]>" + id + "";
                OleDbCommand updateidcmd = new OleDbCommand(updateidsql, conn);
                updateidcmd.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("删除成功");
                conn.Close();
                way = 1;
                mark = 1;
            }
        }
        public void fenlei()
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
    }
}
