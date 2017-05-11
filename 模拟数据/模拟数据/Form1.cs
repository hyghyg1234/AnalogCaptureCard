using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace 模拟数据
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        serial serial = new serial();
        Thread t;
        bool Sendflag = false;
        double[] mean_value = new double[8] { 2.05, 2.05, 2.05, 2.05, 2.05, 2.05, 2.05, 2.05 };
        Random ran = new Random();

        private void timer1_Tick(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == "开始")
            {
                Sendflag = true;
                button2.Text = "停止";
            }
            else
            {
                Sendflag = false;
                button2.Text = "开始";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            serial.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            t = new Thread(DataSend);
            t.Start();
        }

        private void DataSend()
        {
            if (!serial.serialPort1.IsOpen)
            {
                try
                {
                    serial.serialPort1.PortName = Properties.Settings.Default.PortName;
                    serial.serialPort1.BaudRate = Convert.ToInt32(Properties.Settings.Default.BaudRate);
                    serial.serialPort1.Open();
                }
                catch
                { }
            }
            while (true)
            {
                if (Sendflag == false)
                {
                    goto End;
                }
                if (serial.serialPort1.IsOpen)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        serial.serialPort1.Write("CH");
                        serial.serialPort1.Write((i + 1).ToString("00"));
                        serial.serialPort1.Write(" ");
                        serial.serialPort1.Write((mean_value[i] + (double)ran.Next(Convert.ToInt16(textBox4.Text)) / 100).ToString("0.000000"));
                        serial.serialPort1.Write("\r\n");
                        Thread.Sleep(10);
                    }
                    Thread.Sleep(100);
                }
                End:
                    Thread.Sleep(10);
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            t.Abort();
        }
    }
}
