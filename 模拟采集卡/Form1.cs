using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ZedGraph;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using CCWin;

namespace 模拟采集卡
{
    public partial class Form1 : Skin_Mac
    {
        public Form1()
        {
            InitializeComponent();
        }
        public string PortName;
        public double zedGraph_time;
        public double csv_time;
        Thread WriteExcelThread;
        Thread SerialThread;
        int[] CheckArray = new int[8];
        bool StartFlag = false;
        private GraphPane mGraphPane;

        serial serial = new serial();
        CSVHelper csvHelper = new CSVHelper();

        List<System.Windows.Forms.CheckBox> CheckItem = new List<System.Windows.Forms.CheckBox>();     
        private void CheckItem_Add()
        {
            CheckItem.Add(checkBox1);
            CheckItem.Add(checkBox2);
            CheckItem.Add(checkBox3);
            CheckItem.Add(checkBox4);
            CheckItem.Add(checkBox5);
            CheckItem.Add(checkBox6);
            CheckItem.Add(checkBox7);
            CheckItem.Add(checkBox8);
        } 

        //曲线初始化
        #region
        private void init_zedgragh()
        {
            int chartPoint = 200;
            zedGraphControl1.PanModifierKeys = Keys.None;//曲线可以左键拖拽
            zedGraphControl1.ZoomStepFraction = 0.1;//（这是鼠标滚轮缩放的比例大小，值越大缩放就越灵敏）

            zedGraphControl1.IsShowHScrollBar = true;
            mGraphPane = zedGraphControl1.GraphPane;
            mGraphPane.Title.Text = "压力数据";
            //添加两个Y轴，分别显示电压、电流
            mGraphPane.XAxis.Title.Text = "时间";
            mGraphPane.YAxis.Title.Text = "压力值";


            mGraphPane.Y2Axis.IsVisible = false;
            mGraphPane.YAxis.Scale.FontSpec.FontColor = Color.Blue;
            mGraphPane.YAxis.Title.FontSpec.FontColor = Color.Blue;

            mGraphPane.XAxis.Scale.Min = 0;      //X轴最小值0
            mGraphPane.XAxis.Scale.Max = 50;     //时间最大值30分钟
            mGraphPane.XAxis.Scale.MinorStep = 1;//X轴小步长1,也就是小间隔
            mGraphPane.XAxis.Scale.MajorStep = 10;//X轴大步长为5，也就是显示文字的大间隔

            mGraphPane.YAxis.Scale.MinorStep = 1;//X轴小步长1,也就是小间隔
            mGraphPane.YAxis.Scale.MajorStep = 2;//X轴大步长为5，也就是显示文字的大间隔

            try
            {
                mGraphPane.YAxis.Scale.Min = Convert.ToInt16(textBox4.Text);      //电压轴最小值0
                mGraphPane.YAxis.Scale.Max = Convert.ToInt16(textBox3.Text);    //电压最大值
            }
            catch
            {
                MessageBox.Show("参数错误！");
            }
            // Display the Y axis grid lines
            mGraphPane.YAxis.MajorGrid.IsVisible = true;
            mGraphPane.YAxis.MinorGrid.IsVisible = true;

            // Fill the axis background with a color gradient
            mGraphPane.Chart.Fill = new Fill(Color.FromArgb(255, 255, 245), Color.FromArgb(255, 255, 190), 90F);

            mGraphPane.Fill = new Fill(Color.White, Color.FromArgb(220, 255, 255), 45.0f);

            //mGraphPane.CurveList.Clear();
            //LineItem myCurve = mGraphPane.AddCurve("", list1, Color.Blue, SymbolType.None);
            //LineItem myCurve1 = mGraphPane.AddCurve("", list2, Color.Red, SymbolType.None);

            mGraphPane.CurveList.Clear();
            RollingPointPairList[] lists = new RollingPointPairList[8];
            mGraphPane.CurveList.Clear();

            for (int i = 0; i < 8; i++)
            {
                lists[i] = new RollingPointPairList(chartPoint);

                LineItem myCurve = mGraphPane.AddCurve("", lists[i], CheckItem[i].ForeColor, SymbolType.None);
            }
        }
        #endregion

        //datagridview 格式设置
        #region
        private void datagridview_Init()
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;           
            dataGridView1.ReadOnly = true;      //禁止编辑
            dataGridView1.AllowUserToResizeRows = false;//行大小不能调整
            dataGridView1.AllowUserToResizeColumns = false;//行大小不能调整
            //dataGridView1.RowHeadersVisible = false;//行标题隐藏                    
            dataGridView1.DefaultCellStyle.BackColor = Color.Black;
            //dataGridView1.DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("宋体", 12F);
                 
            //dataGridView1添加
            //for (int i = 0; i < 20; i++)
            //{
            //    int index = this.dataGridView1.Rows.Add();
            //}
            //datagridview 格式设置结束
        }
        //dataGridView1数据的DataTable
        private System.Data.DataTable m_GradeTable;

        /// <summary>
        /// 绑定数据
        /// </summary>
        private void BindData()
        {
            //建立一个DataTable并填充数据，然后绑定到DataGridView控件上
            m_GradeTable = new System.Data.DataTable();
            m_GradeTable.Columns.Add("CH1", typeof(string));
            m_GradeTable.Columns.Add("CH2", typeof(string));
            m_GradeTable.Columns.Add("CH3", typeof(string));
            m_GradeTable.Columns.Add("CH4", typeof(string));
            m_GradeTable.Columns.Add("CH5", typeof(string));
            m_GradeTable.Columns.Add("CH6", typeof(string));
            m_GradeTable.Columns.Add("CH7", typeof(string));
            m_GradeTable.Columns.Add("CH8", typeof(string));
            dataGridView1.DataSource = m_GradeTable;
            //禁止排序
            for (int i = 0; i < this.dataGridView1.Columns.Count; i++)
            {
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dataGridView1.Columns[0].DefaultCellStyle.ForeColor = lbDigitalMeter1.ForeColor;
            dataGridView1.Columns[1].DefaultCellStyle.ForeColor = lbDigitalMeter2.ForeColor;
            dataGridView1.Columns[2].DefaultCellStyle.ForeColor = lbDigitalMeter3.ForeColor;
            dataGridView1.Columns[3].DefaultCellStyle.ForeColor = lbDigitalMeter4.ForeColor;
            dataGridView1.Columns[4].DefaultCellStyle.ForeColor = lbDigitalMeter5.ForeColor;
            dataGridView1.Columns[5].DefaultCellStyle.ForeColor = lbDigitalMeter6.ForeColor;
            dataGridView1.Columns[6].DefaultCellStyle.ForeColor = lbDigitalMeter7.ForeColor;
            dataGridView1.Columns[7].DefaultCellStyle.ForeColor = lbDigitalMeter8.ForeColor;
        }
        #endregion

        //参数加载
        #region
        private void parameter_Init()
        {
            textBox4.Text = Properties.Settings.Default.MIN;
            textBox3.Text = Properties.Settings.Default.MAX;
            textBox6.Text = Properties.Settings.Default.RefreshTime;          
            textBox1.Text = Properties.Settings.Default.CsvTime;
            textBox2.Text = Properties.Settings.Default.ZedGraph_time;
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {           
            datagridview_Init();//datagridview初始化
            parameter_Init();   //参数初始化
            BindData();         //数据绑定到表格

            Control.CheckForIllegalCrossThreadCalls = false;
            
            for (int i = 0; i < 8; i++)
            {
                if ((Convert.ToInt16(Properties.Settings.Default.CH_SET) & (0x80 >> i)) > 0)
                {
                    CheckArray[i] = 1;
                }
                else
                {
                    CheckArray[i] = 0;
                }
            }          
            //颜色初始化
            CheckItem_Add();
            CheckItem[0].ForeColor = lbDigitalMeter1.ForeColor;
            CheckItem[1].ForeColor = lbDigitalMeter2.ForeColor;
            CheckItem[2].ForeColor = lbDigitalMeter3.ForeColor;
            CheckItem[3].ForeColor = lbDigitalMeter4.ForeColor;
            CheckItem[4].ForeColor = lbDigitalMeter5.ForeColor;
            CheckItem[5].ForeColor = lbDigitalMeter6.ForeColor;
            CheckItem[6].ForeColor = lbDigitalMeter7.ForeColor;
            CheckItem[7].ForeColor = lbDigitalMeter8.ForeColor;

            init_zedgragh();    //曲线初始化
            
            try
            {
                zedGraph_time = Convert.ToDouble(Properties.Settings.Default.ZedGraph_time);
                csv_time = Convert.ToDouble(Properties.Settings.Default.CsvTime);             
                curveTimer.Interval = (int)(zedGraph_time * 1000);
                dataTimer.Interval = Convert.ToInt16(textBox6.Text);
            }
            catch
            {
                MessageBox.Show("初始化参数错误！");
            }

            SerialThread = new Thread(SerialRead);      //串口数据读取线程
            SerialThread.Start();

            try
            {
                serial.serialPort1.DataReceived += serial.serialPort1_DataReceived;
                serial.serialPort1.PortName = Properties.Settings.Default.PortName;
                serial.serialPort1.BaudRate = 115200;
                serial.serialPort1.Open();
                serial.serialPort1.Write("start\r\n");     //发送上位机启动标志               
            }
            catch
            {
                MessageBox.Show("串口连接错误！");
            }
        }

        string[] SensorString = new string[8];      //8通道电压值字符串格式
        double[] SensorValue = new double[8];       //8通道电压值

        private void SerialRead()
        {           
            int chanel;
            int CH_SET;
            while (true)
            {
                if (StartFlag == false)
                {
                    goto ReadEnd;
                }
                string str = serial.SerialReceivedStr;
                if (str != null)
                {
                    if (str.Contains("Chanel"))
                    {
                        CH_SET = Convert.ToInt16(str.Substring(str.IndexOf("Chanel") + 6, 3));
                        Properties.Settings.Default.CH_SET = CH_SET.ToString("000");
                        Properties.Settings.Default.Save();
                        for (int i = 0; i < 8; i++)
                        {
                            if ((Convert.ToInt16(Properties.Settings.Default.CH_SET) & (0x80 >> i)) > 0)
                            {
                                CheckArray[i] = 1;
                            }
                            else
                            {
                                CheckArray[i] = 0;
                            }
                        }
                    }
                    if (str.Contains("CH"))
                    {
                        chanel = Convert.ToInt16(str.Substring(str.IndexOf("CH") + 2, 2));
                        SensorString[chanel - 1] = str.Substring(str.IndexOf("CH") + 5, 8);
                        SensorValue[chanel - 1] = Convert.ToDouble(SensorString[chanel - 1]);
                    }
                }
                ReadEnd:
                Thread.Sleep(1);
            }
        }

        //曲线更新定时器
        #region
        double time = 0;
        private void timer2_Tick(object sender, EventArgs e)
        {    
            time++;
            LineItem[] curve = new LineItem[32];
            IPointListEdit[] pLists = new IPointListEdit[32];
            //取Graph第一个曲线，也就是第一步:在GraphPane.CurveList集合中查找CurveItem
            for (int i = 0; i < 8; i++)
            {
                curve[i] = zedGraphControl1.GraphPane.CurveList[i] as LineItem;
                pLists[i] = curve[i].Points as IPointListEdit;
            }
            if (CheckItem[0].Checked)
            {
                pLists[0].Add(time, Convert.ToDouble(lbDigitalMeter1.Value));
            }
            if (CheckItem[1].Checked)
            {
                pLists[1].Add(time, Convert.ToDouble(lbDigitalMeter2.Value));
            }
            if (CheckItem[2].Checked)
            {
                pLists[2].Add(time, Convert.ToDouble(lbDigitalMeter3.Value));
            }
            if (CheckItem[3].Checked)
            {
                pLists[3].Add(time, Convert.ToDouble(lbDigitalMeter4.Value));
            }
            if (CheckItem[4].Checked)
            {
                pLists[4].Add(time, Convert.ToDouble(lbDigitalMeter5.Value));
            }
            if (CheckItem[5].Checked)
            {
                pLists[5].Add(time, Convert.ToDouble(lbDigitalMeter6.Value));
            }
            if (CheckItem[6].Checked)
            {
                pLists[6].Add(time, Convert.ToDouble(lbDigitalMeter7.Value));
            }
            if (CheckItem[7].Checked)
            {
                pLists[7].Add(time, Convert.ToDouble(lbDigitalMeter8.Value));
            }

            Scale xScale = zedGraphControl1.GraphPane.XAxis.Scale;
            if (time > xScale.Max - xScale.MajorStep)
            {
                xScale.Max = time + xScale.MajorStep;
                xScale.Min = xScale.Max - 50.0;
                zedGraphControl1.ScrollMaxX = xScale.Max;
                //前面设置点的数目，大于则重现开始从0开始
                if (xScale.Max >= 400)
                {
                    zedGraphControl1.ScrollMinX = 0;
                    zedGraphControl1.ScrollMaxX = 100;
                    zedGraphControl1.GraphPane.XAxis.Scale.Max = 50;
                    zedGraphControl1.GraphPane.XAxis.Scale.Min = 0;
                    for (int i = 0; i < 8; i++)
                    {
                        if (pLists[i] != null)
                            pLists[i].Clear();
                    }
                    time = 0;
                }
            }
            this.zedGraphControl1.Refresh();
        }
        #endregion

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.MAX = textBox3.Text;
            Properties.Settings.Default.Save();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.MIN = textBox4.Text;
            Properties.Settings.Default.Save();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt16(textBox4.Text) >= Convert.ToInt16(textBox3.Text))
                {
                    MessageBox.Show("参数设置错误！");
                    return;
                }
            }
            catch
            {
                MessageBox.Show("参数设置错误！");
                return;
            }
            init_zedgragh();
        }

        //刷新数据的事件
        #region
        int x = 0;  //用来抛弃前几个数据
        private void timer1_Tick(object sender, EventArgs e)
        {           
            x++;
            if (x > 5)
            {
                m_GradeTable.Rows.Add(SensorString);
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
                toolStripStatusLabel1.Text = (dataGridView1.RowCount - 1).ToString() + "行";
            }
            if (CheckArray[0] == 1)
            {
                lbDigitalMeter1.Value = SensorValue[0];
            }
            else
            {
                lbDigitalMeter1.Value = 0;
            }
            if (CheckArray[1] == 1)
            {
                lbDigitalMeter2.Value = SensorValue[1];
            }
            else
            {
                lbDigitalMeter2.Value = 0;
            }
            if (CheckArray[2] == 1)
            {
                lbDigitalMeter3.Value = SensorValue[2];
            }
            else
            {
                lbDigitalMeter3.Value = 0;
            }
            if (CheckArray[3] == 1)
            {
                lbDigitalMeter4.Value = SensorValue[3];
            }
            else
            {
                lbDigitalMeter4.Value = 0;
            }
            if (CheckArray[4] == 1)
            {
                lbDigitalMeter5.Value = SensorValue[4];
            }
            else
            {
                lbDigitalMeter5.Value = 0;
            }
            if (CheckArray[5] == 1)
            {
                lbDigitalMeter6.Value = SensorValue[5];
            }
            else
            {
                lbDigitalMeter6.Value = 0;
            }
            if (CheckArray[6] == 1)
            {
                lbDigitalMeter7.Value = SensorValue[6];
            }
            else
            {
                lbDigitalMeter7.Value = 0;
            }
            if (CheckArray[7] == 1)
            {
                lbDigitalMeter8.Value = SensorValue[7];
            }
            else
            {
                lbDigitalMeter8.Value = 0;
            }         
        }
        #endregion

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            serial.ShowDialog();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (toolStripButton4.Text == "开始采集")
            {
                toolStripButton4.Text = "停止采集";
                toolStripButton4.Image = global::模拟采集卡.Properties.Resources.stop;
                //t = new Thread(WriteExcelData);
                //t.Start();
                curveTimer.Enabled = true;
                StartFlag = true;
                dataTimer.Enabled = true;
                toolStripButton1.Enabled = false;
            }
            else
            {
                toolStripButton4.Text = "开始采集";
                toolStripButton4.Image = global::模拟采集卡.Properties.Resources.start;
                StartFlag = false;
                curveTimer.Enabled = false;
                dataTimer.Enabled = false;
                toolStripButton1.Enabled = true;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            //if (serial.serialPort1.IsOpen)
            //{
            //    serial.serialPort1.Close();
            //}

            //Properties.Settings.Default.PortName = comboBox1.Text;
            //Properties.Settings.Default.CsvTime = textBox1.Text;
            //Properties.Settings.Default.ZedGraph_time = textBox2.Text;
            //Properties.Settings.Default.RefreshTime = textBox6.Text;
            //Properties.Settings.Default.Save();

            //try
            //{
            //    zedGraph_time = Convert.ToDouble(textBox2.Text);
            //    csv_time = Convert.ToInt16(textBox1.Text);

            //    timer1.Interval = Convert.ToInt16(textBox6.Text);
            //    timer2.Interval = (int)(zedGraph_time * 1000);
            //}
            //catch
            //{
            //    MessageBox.Show("请填写正确参数！");
            //    return;
            //}
            //try
            //{
            //    serialPort1.PortName = comboBox1.Text;
            //    serialPort1.Open();
            //    serialPort1.Write("start\r\n");     //发送上位机启动标志
            //    MessageBox.Show("设置完成！");
            //}
            //catch
            //{
            //    MessageBox.Show("串口连接错误！");
            //}    
        }

        /// <summary>
        /// 事件函数
        /// </summary>
        /// <param name="s"></param>
        public void SetSerialWriteEvent(string s)
        {
            if (!serial.serialPort1.IsOpen)
            {
                serial.serialPort1.Open();
            }
            serial.serialPort1.Write(s);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Setting Set = new Setting();
            string CH_SET;
            Set.SerialWriteEvent += SetSerialWriteEvent;    //事件添加函数
            if (Set.ShowDialog() == DialogResult.OK)
            {
                CH_SET = Properties.Settings.Default.CH_SET;
                serial.serialPort1.Write("CH");
                serial.serialPort1.Write(CH_SET);
                serial.serialPort1.Write(" SET");
                serial.serialPort1.Write(Properties.Settings.Default.ADS1256_DRATE);
                serial.serialPort1.Write("end\r\n");

                for (int i = 0; i < 8; i++)
                {
                    if ((Convert.ToInt16(Properties.Settings.Default.CH_SET) & (0x80 >> i)) > 0)
                    {
                        CheckArray[i] = 1;
                    }
                    else
                    {
                        CheckArray[i] = 0;
                    }
                }
            }
        }

        //右击清除
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //while (dataGridView1.RowCount > 1)
            //{
            //    m_GradeTable.Rows[0].Delete();
            //}
            m_GradeTable.Clear();
            toolStripStatusLabel1.Text = "0行";
        }
        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog FileSave = new SaveFileDialog();
            FileSave.Title = "保存EXECL文件";
            FileSave.Filter = "CSV文件(*.csv) |*.csv | 所有文件(*.*) |*.*";
            FileSave.FilterIndex = 1;
            if (FileSave.ShowDialog() == DialogResult.OK)
            {
                string FileName = FileSave.FileName;
                if (File.Exists(FileName))
                {
                    File.Delete(FileName);
                }
                csvHelper.DataTableToCSV(m_GradeTable, FileName);
                MessageBox.Show(this, "保存CSV成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "CSV文件|*.CSV";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog1.FileName;
                System.Data.DataTable newDataTable;
                newDataTable = csvHelper.CSVToDataTable(fileName);
                m_GradeTable.Clear();    
                foreach (DataRow row in newDataTable.Rows)
                {
                    m_GradeTable.ImportRow(row);
                }
                toolStripStatusLabel1.Text = (dataGridView1.RowCount - 1).ToString() + "行";
                MessageBox.Show("成功显示CSV数据！");
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "CSV文件|*.CSV";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog1.FileName;
                System.Data.DataTable newDataTable;
                newDataTable = csvHelper.CSVToDataTable(fileName);
                m_GradeTable.Clear();
                foreach (DataRow row in newDataTable.Rows)
                {
                    m_GradeTable.ImportRow(row);
                }
                toolStripStatusLabel1.Text = (dataGridView1.RowCount - 1).ToString() + "行";
                MessageBox.Show("成功显示CSV数据！");
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            SaveFileDialog FileSave = new SaveFileDialog();
            FileSave.Title = "保存EXECL文件";
            FileSave.Filter = "CSV文件(*.csv) |*.csv | 所有文件(*.*) |*.*";
            FileSave.FilterIndex = 1;
            if (FileSave.ShowDialog() == DialogResult.OK)
            {
                string FileName = FileSave.FileName;
                if (File.Exists(FileName))
                {
                    File.Delete(FileName);
                }
                csvHelper.DataTableToCSV(m_GradeTable, FileName);
                MessageBox.Show(this, "保存CSV成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //关闭程序
        private void Close_Form1()
        {
            SerialThread.Abort();          
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Close_Form1();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            Close_Form1();
            this.Close();
        }       
    }
}
