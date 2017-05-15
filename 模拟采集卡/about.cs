using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 模拟采集卡
{
    public partial class about : Form
    {
        public about()
        {
            InitializeComponent();
        }

        private void linkLabel1_Click(object sender, EventArgs e)
        {
            //调用系统默认的浏览器   
            System.Diagnostics.Process.Start(linkLabel1.Text);
        }
    }
}
