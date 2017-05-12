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
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }
        //定义带参数的委托与事件
        public delegate void Mydelegate(string s);
        public event Mydelegate SerialWriteEvent;

        int chanel;
        int ADS1256_DRATE;
        int[] ADS1256_DRATE_BUF = new int[16] { 0x03, 0x13 , 0x23 , 0x33 , 0x43 , 0x53 ,
            0x63 , 0x72 , 0x82 , 0x92 , 0xA1 , 0xB0 , 0xC0, 0xD0 , 0xE0 , 0xF0 };           //ADS1256转换速度写入值

        List<CheckBox> checkbox_item = new List<CheckBox>();

        List<RadioButton> RadioButtonItem = new List<RadioButton>();

        serial serial = new serial();

        private void checkbox_Add()
        {
            checkbox_item.Add(checkBox1);
            checkbox_item.Add(checkBox2);
            checkbox_item.Add(checkBox3);
            checkbox_item.Add(checkBox4);
            checkbox_item.Add(checkBox5);
            checkbox_item.Add(checkBox6);
            checkbox_item.Add(checkBox7);
            checkbox_item.Add(checkBox8);
        }

        private void RadioButtonItem_Add()
        {
            RadioButtonItem.Add(raBtton0);
            RadioButtonItem.Add(raBtton1);
            RadioButtonItem.Add(raBtton2);
            RadioButtonItem.Add(raBtton3);
            RadioButtonItem.Add(raBtton4);
            RadioButtonItem.Add(raBtton5);
            RadioButtonItem.Add(raBtton6);
            RadioButtonItem.Add(raBtton7);
            RadioButtonItem.Add(raBtton8);
            RadioButtonItem.Add(raBtton9);
            RadioButtonItem.Add(raBtton10);
            RadioButtonItem.Add(raBtton11);
            RadioButtonItem.Add(raBtton12);
            RadioButtonItem.Add(raBtton13);
            RadioButtonItem.Add(raBtton14);
            RadioButtonItem.Add(raBtton15);
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            checkbox_Add();
            RadioButtonItem_Add();

            for (int i = 0; i < 16; i++)
            {
                if (Convert.ToInt16(Properties.Settings.Default.ADS1256_DRATE) == ADS1256_DRATE_BUF[i])
                {
                    RadioButtonItem[i].Checked = true;
                }
            }

            for (int i = 0; i < 8; i++)
            {
                if ((Convert.ToInt16(Properties.Settings.Default.CH_SET) & (0x80 >> i)) > 0)
                {
                    checkbox_item[i].Checked = true;
                }
            }              
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
            {
                checkbox_item[i].Checked = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
            {
                checkbox_item[i].Checked = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
            {
                if (checkbox_item[i].Checked == true)
                {
                    checkbox_item[i].Checked = false;
                }
                else
                {
                    checkbox_item[i].Checked = true;
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate1\r\n");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate2\r\n");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate3\r\n");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate4\r\n");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate5\r\n");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate6\r\n");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate7\r\n");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SerialWriteEvent("Calibrate8\r\n");
        }

        private void button14_Click(object sender, EventArgs e)
        {
            int[] CheckArray = new int[8];
            for (int i = 0; i < 8; i++)
            {
                if (checkbox_item[i].Checked == true)
                {
                    CheckArray[i] = 1;
                }
                else
                {
                    CheckArray[i] = 0;
                }
            }

            chanel = CheckArray[0] * 128 + CheckArray[1] * 64 + CheckArray[2] * 32 + CheckArray[3] * 16 +
                CheckArray[4] * 8 + CheckArray[5] * 4 + CheckArray[6] * 2 + CheckArray[7] * 1;

            for (int i = 0; i < 16; i++)
            {
                if (RadioButtonItem[i].Checked == true)
                {
                    ADS1256_DRATE = ADS1256_DRATE_BUF[i];
                }
            }

            Properties.Settings.Default.ADS1256_DRATE = ADS1256_DRATE.ToString("000");
            Properties.Settings.Default.CH_SET = chanel.ToString("000");
            Properties.Settings.Default.Save();
            DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.OK;
        }
    }
}
