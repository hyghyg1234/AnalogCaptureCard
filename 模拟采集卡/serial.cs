using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 模拟采集卡
{
    public partial class serial : Form
    {
        public serial()
        {
            InitializeComponent();
            this.comboBox2.Items.AddRange(new object[] {
            "4800",
            "9600",
            "14400",
            "19200",
            "38400",
            "43000",
            "57600",
            "76800",
            "115200",
            "128000",
            "230400",
            "256000",
            "460800",
            "921600",
            "1382400"
            });
        }
        public SerialPort serialPort1 = new SerialPort();

        /*串口端口扫描功能*/
        private void SearchAndAddSerialToComboBox(ComboBox Mybox)
        {
            Mybox.Items.Clear();
            String[] Portname = SerialPort.GetPortNames();
            foreach (string str in Portname)
            {
                Mybox.Items.Add(str);
            }
        }

        private void serial_Load(object sender, EventArgs e)
        {
            SearchAndAddSerialToComboBox(comboBox1);      
            comboBox1.Items.Add(Properties.Settings.Default.PortName);
            comboBox1.Text = Properties.Settings.Default.PortName;
            comboBox2.Text = Properties.Settings.Default.BaudRate;         
        }

        public string SerialReceivedStr = null;

        public void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {          
            try
            {
                SerialReceivedStr = serialPort1.ReadTo("\r\n");
            }
            catch
            { }
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            SearchAndAddSerialToComboBox(comboBox1);
        }

        private void serial_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.PortName = comboBox1.Text;
            Properties.Settings.Default.BaudRate = comboBox2.Text;
            Properties.Settings.Default.Save();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    serialPort1.Close();
                }
                serialPort1.PortName = comboBox1.Text;
                serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text, 10);
                serialPort1.Open();
            }
            catch
            {
                MessageBox.Show("串口开启错误，请检查！");
            }
            this.Close();
        }
    }
}
