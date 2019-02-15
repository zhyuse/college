using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace 财务销售
{
    public partial class mainpage : Form
    {
        public static int xiaotui = 0, shuaxin = 0;//判断是销售还是退货，默认0销售,1退货
        public mainpage()
        {
            InitializeComponent();
        }

        private void mainpage_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult.Cancel == MessageBox.Show("确认退出？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information))
            {
                e.Cancel = true;
            }
        }

        private void mainpage_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            guanli form = new guanli();
            form.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dayin form = new dayin();
            form.Show();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(shuaxin==1)
            {
                button2.PerformClick();
                shuaxin++;
            }
        }
    }
}
