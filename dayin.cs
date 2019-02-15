using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
////using System.Linq;
using System.Text;
////using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace 财务销售
{
    public partial class dayin : Form
    {
        public static double price1, price2;
        public static string bianhao,product,danwei,beizhu;
        public static string productdel;
        public static double numberdel, danjiadel,pricedel;
        public static int mark = 0, way = 1, m = 0, m2 = 0, page, page2, nowpage = 1, nowpage2 = 1, allnum, allnum2;
        public static double allmoney;//m-已查询的行数
        public static int fanye = 0;
        public dayin()
        {
            InitializeComponent();
            this.printDocument1.OriginAtMargins = true;//启用页边距
            this.pageSetupDialog1.EnableMetric = true; //以毫米为单位
        }

        private void dayin_Load(object sender, EventArgs e)
        {
            // DataGridView1.EnableHeadersVisualStyles = false;  设置行标题颜色需要设置这个
            //AllowUserToOrderColumns=true;//允许用户拖动列
            //dataGridView1.AllowUserToResizeColumns = false;//锁定列宽
            dataGridView1.AllowUserToResizeRows = false;//锁定行高
            //dataGridView2.AllowUserToResizeColumns = false;//锁定列宽
            dataGridView2.AllowUserToResizeRows = false;//锁定行高
            //标题行高设置 1修改ColumnHeadersHeader 设置为你想要的高度，比如20；但这时候自动变回来。2修改ColumnHeadersHeaderSize属性为 EnableResizing，不要为AutoSize。
            //行高dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;dataGridView1.RowTemplate.Height = 100;
            showall();
            label11.Text = DateTime.Now.ToString("yyyy年MM月dd日");
            label10.Visible = false;
            上一页.Enabled = false;
            button8.Enabled = false;
            show2();
            label3.Text = "总共有" + page + "页，当前第" + nowpage + "页";
            label8.Text = "第1页";
            if (allnum <= 17)
                下一页.Enabled = false;
            if (allnum2 <= 16)
                button9.Enabled = false;
            fenlei();
            kehu();
            dataGridView2.Columns[0].Width = 100;
            dataGridView2.Columns[1].Width = 300;
            dataGridView2.Columns[4].Width = 100;
            dataGridView1.Columns[0].Width = 300;
            dataGridView1.Columns[2].Width = 65;
            if (mainpage.xiaotui == 0)
                label4.Text = "平发电器销售单";
            else
                label4.Text = "平发电器退货单";

        }
        public void showall()
        {
            dataGridView2.AllowUserToAddRows = false;//删掉最后一行
            dataGridView2.RowHeadersVisible = false;//删掉左边一列
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string sqlnum = @"select count(*) from [" + Form1.id + "]";
            OleDbCommand cmdnum = new OleDbCommand(sqlnum, conn);
            allnum = Convert.ToInt32(cmdnum.ExecuteScalar());
            if (allnum % 17 == 0)
                page = allnum / 17;
            else
                page = allnum / 17 + 1;
            string sql;
            //OleDbDataAdapter da = new OleDbDataAdapter("SELECT * from ["+Form1.id+"] ", conn);
            if (m == 0)
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] ORDER BY [beizhu], [product]";
            else
                sql = "SELECT top 17 [bianhao] as 编号, [product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] where id not in(select top " + m + " id from [" + Form1.id + "] order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        public void showfenlei()
        {
            dataGridView2.AllowUserToAddRows = false;
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
                sql = "SELECT top 17 [bianhao] as 编号,[product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] where [beizhu]='" + comboBox1.Text + "' ORDER BY [beizhu], [product]";
            else
                sql = "SELECT top 17 [bianhao] as 编号,[product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from[" + Form1.id + "] where [beizhu]='" + comboBox1.Text + "' and id not in(select top " + m + " id from [" + Form1.id + "] where [beizhu]='" + comboBox1.Text + "' order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        public void showabout()
        {
            dataGridView2.AllowUserToAddRows = false;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string text = "%" + textBox2.Text + "%";//暂时用这种方法来实现模糊查询
            string sqlnum;
            sqlnum = @"select count(*) from [" + Form1.id + "] where [product] like '" + text + "' ";
            OleDbCommand cmdnum = new OleDbCommand(sqlnum, conn);
            allnum = Convert.ToInt32(cmdnum.ExecuteScalar());
            if (allnum % 17 == 0)
                page = allnum / 17;
            else
                page = allnum / 17 + 1;
            string sql;
            if (m != 0)
                sql = "SELECT top 17 [bianhao] as 编号,[product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [product] like '" + text + "' and id not in(select top " + m + " id from [" + Form1.id + "] where [product] like '" + text + "' order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            else
                sql = "SELECT top 17 [bianhao] as 编号,[product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [product] like '" + text + "' ORDER BY [beizhu], [product]";
            //access模糊查询要用*
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        public void showbianhao()
        {
            dataGridView2.AllowUserToAddRows = false;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string text = "%" + textBox1.Text + "%";//暂时用这种方法来实现模糊查询
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
                sql = "SELECT top 17 [bianhao] as 编号,[product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [bianhao] like '" + text + "' and id not in(select top " + m + " id from [" + Form1.id + "] where [bianhao] like '" + text + "' order by [beizhu], [product]) ORDER BY [beizhu], [product]";
            else
                sql = "SELECT top 17 [bianhao] as 编号,[product] as 产品,[danwei] as 单位,[price2] as 售价,[beizhu] as 分类 from [" + Form1.id + "] where [bianhao] like '" + text + "' ORDER BY [beizhu], [product]";
            //access模糊查询要用*
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "" + Form1.id + "");
            dataGridView2.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
        public void show2()
        {
            dataGridView1.AllowUserToAddRows = false;//删掉最后一行
            dataGridView1.RowHeadersVisible = false;//删掉左边一列
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            string sqlnum = @"select count(*) from [dayin]";
            OleDbCommand cmdnum = new OleDbCommand(sqlnum, conn);
            allnum2 = Convert.ToInt32(cmdnum.ExecuteScalar());
            if (allnum2 % 16 == 0)
                page2 = allnum2 / 16;
            else
                page2 = allnum2 / 16 + 1;
            string sql;
            if(mainpage.xiaotui==0)
            {
                if (m2 == 0)
                    sql = "SELECT top 16 [product] as 产品,[danwei] as 单位,[number] as 数量,[danjia] as 单价,[price] as 金额 from[dayin] ORDER BY [id]";
                else
                    sql = "SELECT top 16 [product] as 产品,[danwei] as 单位,[number] as 数量,[danjia] as 单价,[price] as 金额 from[dayin] where id not in(select top " + m2 + " id from [dayin] order by [id]) ORDER BY [id]";
                //OleDbDataAdapter da = new OleDbDataAdapter("SELECT * from ["+Form1.id+"] ", conn);
                //sql = "SELECT  [product] as 产品,[danwei] as 单位,[number] as 数量,[danjia] as 单价,[price] as 金额 from[dayin] ORDER BY [id]";
            }
            else 
            {
                if (m2 == 0)
                    sql = "SELECT top 16 [product] as 产品,[number] as 数量,[danwei] as 单位,[danjia] as 单价,[price] as 金额 from[dayin] ORDER BY [id]";
                else
                    sql = "SELECT top 16 [product] as 产品,[number] as 数量,[danwei] as 单位,[danjia] as 单价,[price] as 金额 from[dayin] where id not in(select top " + m2 + " id from [dayin] order by [id]) ORDER BY [id]";
            }
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "dayin");
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
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
        public void kehu()
        {
            string sql = @"select [name] from [kehu] ";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds, "kehu");
            comboBox2.DataSource = ds;
            comboBox2.DisplayMember = "kehu.name";
            conn.Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            pageSetupDialog1.Document = printDocument1;
            this.pageSetupDialog1.ShowDialog();
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {//return;
                DataGridViewRow dgvr = dataGridView2.CurrentRow;
                bianhao = dgvr.Cells["编号"].Value.ToString();
                product = dgvr.Cells["产品"].Value.ToString();
                danwei = dgvr.Cells["单位"].Value.ToString();
                price2 = Convert.ToDouble(dgvr.Cells["售价"].Value);
                beizhu = dgvr.Cells["分类"].Value.ToString();
            }
        }

        private void 上一页_Click(object sender, EventArgs e)
        {
            m -= 17;
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
            label3.Text = "总共有" + page + "页，当前第" + nowpage + "页";
        }

        private void button7_Click(object sender, EventArgs e)//分类查询
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

        private void button6_Click(object sender, EventArgs e)//模糊查询
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

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {//return;
                DataGridViewRow dgvr = dataGridView2.CurrentRow;
                bianhao = dgvr.Cells["编号"].Value.ToString();
                product = dgvr.Cells["产品"].Value.ToString();
                danwei = dgvr.Cells["单位"].Value.ToString();
                price2 = Convert.ToDouble(dgvr.Cells["售价"].Value);
                beizhu = dgvr.Cells["分类"].Value.ToString();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(mark==1)
            {
                show2();
                mark--;
                if (fanye == 1)
                    if (allnum2 % 16 == 1)
                    {
                        button9.Enabled = true;
                        label9.Text = "合计:                                                       ¥";
                    }
                if (fanye == -1)
                    if (allnum2 % 16 == 0)
                        button9.Enabled = false;
            }
            if (page2 == nowpage2)
            {
                calute();
                label9.Text = "合计:                                                       ¥" + allmoney + " ";
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
           mainpage.xiaotui = 0;
            show2();
            label4.Text = "平发电器销售单";
            mainpage.shuaxin = 1;
            this.Close();
        }

        private void button13_Click(object sender, EventArgs e)
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

        private void button12_Click(object sender, EventArgs e)
        {
            mainpage.xiaotui = 1;
            show2();
            label4.Text = "平发电器退货单";
            mainpage.shuaxin = 1;
            this.Close();
        }



        private void button10_Click(object sender, EventArgs e)//移除全部
        {
            string sql = @"delete from [dayin]  ";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            MessageBox.Show("删除成功");
            conn.Close();
            mark = 1;
            button9.Enabled = false;
            button8.Enabled = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            m2 -= 16;
            show2();
            if (m2 == 0)
                button8.Enabled = false;
            if (m2 < allnum2 - 16)
                button9.Enabled = true;
            nowpage2--;
            label8.Text = "第" + nowpage2 + "页";
            if (page2 > nowpage2)
                label9.Text = "合计:                                                          ¥";
        }
        public void calute()
        {
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            OleDbCommand cmd = new OleDbCommand(@"select sum(price) from [dayin] ", conn);
            allmoney = Convert.ToDouble(cmd.ExecuteScalar());
        }
        private void button9_Click(object sender, EventArgs e)
        {
            m2 += 16;
            show2();
            if (m2 >= allnum2 - 16)
                button9.Enabled = false;//下一页
            if (m2 > 0)
                button8.Enabled = true;//上一页
            nowpage2++;
            label8.Text = "第" + nowpage2 + "页";
            if (page2 == nowpage2)
            {
                calute();
                label9.Text = "合计:                                                       ¥" + allmoney + " ";
            }
        }

        private void button5_Click(object sender, EventArgs e)//移除一个
        {
            int id;
            string idsql = @"select id from [dayin] where product='" + productdel + "' ";
            string sql = @"delete from [dayin] where product='" + productdel + "' ";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            OleDbCommand idcmd = new OleDbCommand(idsql, conn);
            id = Convert.ToInt32(idcmd.ExecuteScalar());
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.ExecuteNonQuery();
            string updateidsql = @"update [dayin] set [id]=[id]-1 where [id]>" + id + "";
            OleDbCommand updateidcmd = new OleDbCommand(updateidsql, conn);
            updateidcmd.ExecuteNonQuery();
            conn.Close();
            fanye = -1;
            MessageBox.Show("删除成功");
            conn.Close();
            mark = 1;
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {//return;
                DataGridViewRow dgvr = dataGridView1.CurrentRow;
                productdel = dgvr.Cells["产品"].Value.ToString();
                numberdel = Convert.ToInt32(dgvr.Cells["数量"].Value);
                danjiadel = Convert.ToInt32(dgvr.Cells["单价"].Value);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.Document = printDocument1;
            this.printPreviewDialog1.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text == string.Empty)
            {
                MessageBox.Show("请输入客户名称");
                return;
            }
            kehupd();
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.White;
            while (button8.Enabled == true)
                button8.PerformClick();
            if (this.printDialog1.ShowDialog() == DialogResult.OK)
            {
                this.printDocument1.Print();
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            ////打印内容 为 整个Form
            // Image myFormImage;
            // myFormImage = new Bitmap(this.Width, this.Height);
            // Graphics g = Graphics.FromImage(myFormImage);
            //g.CopyFromScreen(this.Location.X, this.Location.Y, 0, 0, this.Size);
            //e.Graphics.DrawImage(myFormImage, 0, 0);
            label10.Text = comboBox2.Text;
            label10.Visible = true;
            comboBox2.Visible = false;
            if (page2 > nowpage2)
            {
                Bitmap _NewBitmap = new Bitmap(groupBox1.Width, groupBox1.Height);
               groupBox1.DrawToBitmap(_NewBitmap, new Rectangle(0, 0, _NewBitmap.Width, _NewBitmap.Height));
                e.Graphics.DrawImage(_NewBitmap, 0, 0, _NewBitmap.Width, _NewBitmap.Height);
                e.HasMorePages = true; button9.PerformClick();        
            }//此时，系统会重新调用printDocument1_PrintPage方法
            else 
            { //e.HasMorePages = false;
                Bitmap _NewBitmap = new Bitmap(groupBox1.Width, groupBox1.Height);
                groupBox1.DrawToBitmap(_NewBitmap, new Rectangle(0, 0, _NewBitmap.Width, _NewBitmap.Height));
                e.Graphics.DrawImage(_NewBitmap, 0, 0, _NewBitmap.Width, _NewBitmap.Height);
                e.HasMorePages = false;
                label10.Visible = false;
                comboBox2.Visible = true;
            } //此时，系统不会再调用printDocument1_PrintPage方法    
        }
        
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox1_Paint(object sender, PaintEventArgs e)
        {
           
        }

        private void button4_Click(object sender, EventArgs e)//添加窗体
        {
            if (pd())
            {
                dayin2 form = new dayin2();
                form.Show();
            }
            else
                MessageBox.Show("销售单中已存在相同产品");
        }
        public bool pd()//检测product和beizhu是否会出现相同值
        {
            string sql = @"select * from [dayin] where [product]='" + product + "' ";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
                return false;
            else
                return true;
        }
        public void kehupd()
        {
            int id;
            string sqlinsert,sqlid;
            string sql = @"select * from [kehu] where [name]='" + comboBox2.Text + "' ";
            sqlid = @"select count(*) from [kehu]";
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb");
            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                OleDbCommand cmdid = new OleDbCommand(sqlid, conn);
                id = Convert.ToInt32(cmdid.ExecuteScalar())+1;
                sqlinsert = @"insert into [kehu](id,name) values(" + id + ",'" + comboBox2.Text + "')";
                OleDbCommand cmdinsert = new OleDbCommand(sqlinsert, conn);
                cmdinsert.ExecuteNonQuery();
            }
            conn.Close();
        }
    }
}
