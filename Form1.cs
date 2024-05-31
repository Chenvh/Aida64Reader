using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.MemoryMappedFiles;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Data.SqlClient;

namespace Aida64Reader
{
    public partial class Form1 : Form
    {

        public struct AIDA64Item
        {
            internal string id;
            internal string value;
            internal string label;
        }

        private bool FirstItemAddFlag = true;
        private String[] args;
        private IList<string> portList = new List<string>();
        private string portName;
        private const string defaultNullPort = "";
        public string SendString = string.Empty;

        System.Timers.Timer t = new System.Timers.Timer();

        public Form1()
        {
            InitializeComponent();
            InitTimer();
            InitListView();
        }
        private void InitListView()
        {
            lv.GridLines = true;      // 表格是否显示网格线
            lv.FullRowSelect = true;  // 是否选中整行
            lv.View = View.Details;   // 设置显示方式
            lv.Scrollable = true;     // 是否自动显示滚动条
            lv.MultiSelect = false;   // 是否可以选择多行
                                      //lv.CheckBoxes = true;     // 显示复选框

            // group
            ListViewGroup groupsys = new ListViewGroup("系统");
            ListViewGroup grouptemp = new ListViewGroup("温度");
            ListViewGroup groupfan = new ListViewGroup("风扇");
            ListViewGroup groupduty = new ListViewGroup("duty");
            ListViewGroup groupfvol = new ListViewGroup("电压");
            ListViewGroup grouppwr = new ListViewGroup("功耗");
            lv.Groups.Add(groupsys);
            lv.Groups.Add(grouptemp);
            lv.Groups.Add(groupfan);
            lv.Groups.Add(groupduty);
            lv.Groups.Add(groupfvol);
            lv.Groups.Add(grouppwr);
            lv.ShowGroups = true;
            lv.Columns.Clear();
            lv.Columns.Add("名称", 160, HorizontalAlignment.Center);
            lv.Columns.Add("值", 80, HorizontalAlignment.Center);
        }



        /*
         * 初始化定时器
         */
        private void InitTimer()
        {
            textSendInterval.Value = 500;//默认值为1000毫秒
            t.Interval = 500;//实例化Timer类，设置默认间隔时间为1000毫秒
            t.Elapsed += ReadInfoAndSend;//到达时间的时候执行事件
            t.AutoReset = true;//设置是执行一次（false）还是一直执行(true
            t.SynchronizingObject = this;
            t.Start();
        }

        private bool IsPortAvailable()
        {
            return (portName != null && !portName.Equals(""));//没有找到串口的话直接返回
        }

        //定时任务，用于将从AIDA64获取到的信息发送到指定的串口
        private void ReadInfoAndSend(object source, System.Timers.ElapsedEventArgs e)
        {
            string SerialSendData = string.Empty;
            if (!serialPort1.IsOpen)
            {
                return;
            }

            try
            {
                string tempString = string.Empty;
                tempString += "<AIDA64>";

                MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("AIDA64_SensorValues");
                MemoryMappedViewAccessor accessor = mmf.CreateViewAccessor();
                tempString = tempString + "";
                for (int i = 0; i < accessor.Capacity; i++)
                {
                    tempString = tempString + (char)accessor.ReadByte(i);
                }
                tempString = tempString.Replace("\0", "");
                tempString = tempString + "";
                accessor.Dispose();
                mmf.Dispose();

                tempString += "</AIDA64>";
                XDocument aidaXML = XDocument.Parse(tempString);
                var sysElements = aidaXML.Element("AIDA64").Elements("sys");
                var tempElements = aidaXML.Element("AIDA64").Elements("temp");
                var fanElements = aidaXML.Element("AIDA64").Elements("fan");
                var dutyElements = aidaXML.Element("AIDA64").Elements("duty");
                var voltElements = aidaXML.Element("AIDA64").Elements("volt");
                var pwrElements = aidaXML.Element("AIDA64").Elements("pwr");

                List<AIDA64Item> items = new List<AIDA64Item>();


                lv.BeginUpdate();
                lv.Items.Clear();

                lv.Items[1].SubItems.Clear();
                foreach (var i in sysElements)
                {
                    //Console.WriteLine(i.Element("id").Value + "\t" + i.Element("label").Value + "\t" + i.Element("value").Value);
                    AIDA64Item item = new AIDA64Item();
                    var lvitem = new ListViewItem();

                    lvitem.SubItems.Clear();
                    lvitem.Group = lv.Groups[0];
                    if (FirstItemAddFlag) lvitem.SubItems[0].Text = i.Element("label").Value;
                    lvitem.SubItems.Add(i.Element("value").Value);
                    lv.Items.Add(lvitem);
                    //listBox3.Items.Add(tempitem);
                    //内存数据根据注册表顺序，从上往下读取
                    switch (i.Element("id").Value)
                    {
                        case "SCPUCLK":  //CPU时钟频率 {"SCPUCLK":"4100",}
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "SGPU1CLK":  //GPU时钟频率 {"SCPUCLK":"4100",}
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "SCPUUTI":  //CPU使用率
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "SGPU1UTI":  //CPU使用率
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "SMEMUTI": //内存使用率
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "SUSEDMEM": //已使用内存
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "SFREEMEM": //可用内存
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        default:
                            break;
                    }
                    if (item.id != null)
                    {
                        items.Add(item);
                    }

                }
                foreach (var i in tempElements)
                {
                    //Console.WriteLine(i.Element("id").Value + "\t" + i.Element("label").Value + "\t" + i.Element("value").Value);
                    AIDA64Item item = new AIDA64Item();
                    var lvitem = new ListViewItem();

                    lvitem.SubItems.Clear();
                    lvitem.Group = lv.Groups[1];
                    if (FirstItemAddFlag) lvitem.SubItems[0].Text = i.Element("label").Value;
                    lvitem.SubItems.Add(i.Element("value").Value);
                    lv.Items.Add(lvitem);
                    switch (i.Element("id").Value)
                    {
                        case "TCPU":   //CPU温度
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "TGPU1":   //GPU温度
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        default:
                            break;
                    }
                    if (item.id != null)
                    {
                        items.Add(item);
                    }

                }
                foreach (var i in voltElements)
                {
                    //Console.WriteLine(i.Element("id").Value + "\t" + i.Element("label").Value + "\t" + i.Element("value").Value);
                    AIDA64Item item = new AIDA64Item();
                    var lvitem = new ListViewItem();

                    lvitem.SubItems.Clear();
                    lvitem.Group = lv.Groups[4];
                    if (FirstItemAddFlag) lvitem.SubItems[0].Text = i.Element("label").Value;
                    lvitem.SubItems.Add(i.Element("value").Value);
                    lv.Items.Add(lvitem);
                    switch (i.Element("id").Value)
                    {
                        case "VCPU":   //CPU电压
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        case "VGPU1":   //GPU电压
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        default:
                            break;
                    }
                    if (item.id != null)
                    {
                        items.Add(item);
                    }


                }
                foreach (var i in fanElements)
                {
                    //Console.WriteLine(i.Element("id").Value + "\t" + i.Element("label").Value + "\t" + i.Element("value").Value);
                    AIDA64Item item = new AIDA64Item();
                    var lvitem = new ListViewItem();

                    lvitem.SubItems.Clear();
                    lvitem.Group = lv.Groups[3];
                    if (FirstItemAddFlag) lvitem.SubItems[0].Text = i.Element("label").Value;
                    lvitem.SubItems.Add(i.Element("value").Value);
                    lv.Items.Add(lvitem);
                    switch (i.Element("id").Value)
                    {
                        case "FCPU":  //CPU散热风扇
                            item.id = i.Element("id").Value;
                            item.value = i.Element("value").Value;
                            break;
                        default:
                            break;
                    }
                    if (item.id != null)
                    {
                        items.Add(item);
                    }

                }
                FirstItemAddFlag = false;
                lv.EndUpdate();

                //string updateCmd = string.Empty + "update";
                string updateCmd = string.Empty;
                //数据打包成json格式
                foreach (var i in items)
                {
                    updateCmd += ("\"" + i.id + "\":\"" + i.value + "\",");
                }
                updateCmd = updateCmd.Substring(0, (updateCmd.Length - 1)); //移出最后一个逗号
                updateCmd = "{" + updateCmd + "}";
                updateCmd += "\r\n";//下位机以‘\r’为一行的结束
                Console.Write(updateCmd);
                textBox_send.Text = updateCmd;//这里的this指针在设置上面的t.SynchronizingObject = this;
                serialPort1.Write(updateCmd);
                //writeDataToPort(updateCmd);
            }
            catch (Exception ex)
            {
                handleException(ex);
            }

        }
        private void writeDataToPort(string dataStr)
        {
            //SerialPort serialPort1 = new SerialPort(portName, 9600, Parity.None, 8, StopBits.One); //先到设备管理器里找串口对应的端口
            serialPort1.Open();
            serialPort1.Write(dataStr);
            serialPort1.Close();
        }
        private void handleException(Exception ex)
        {
            Console.Write(ex.Message + ",可能已经拔下串口设备或者串口被占用或者发生其他异常");
            //MessageBox.Show(ex.Message + ",可能已经拔下串口设备或者串口被占用或者发生其他异常", "发生异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //this.portName = null;
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            textSendInterval.Value = 1000;//默认值为1000毫秒

            //批量添加波特率列表
            string[] baud = { "4800", "9600", "43000", "56000", "57600", "115200", "128000", "230400", "256000", "460800" };
            comboBox2.Items.AddRange(baud);


            //设置默认值
            //  comboBox1.Text = "COM1";
            comboBox2.Text = "9600";
            comboBox3.Text = "8";
            comboBox4.Text = "None";
            comboBox5.Text = "1";

            //获取电脑当前可用串口并添加到选项列表中
            object[] comitems = System.IO.Ports.SerialPort.GetPortNames();
            comboBox1.Items.AddRange(comitems);
            comboBox1.Text = comitems.First().ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //将可能产生异常的代码放置在try块中
                //根据当前串口属性来判断是否打开
                if (serialPort1.IsOpen)
                {
                    //串口已经处于打开状态
                    serialPort1.Close();    //关闭串口
                    openport_bt.Text = "打开串口";
                    openport_bt.BackColor = Color.ForestGreen;
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = true;
                    comboBox5.Enabled = true;
                    textBox_receive.Text = "";  //清空接收区
                    //textBox_send.Text = "";     //清空发送区
                }
                else
                {
                    //串口已经处于关闭状态，则设置好串口属性后打开
                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    comboBox3.Enabled = false;
                    comboBox4.Enabled = false;
                    comboBox5.Enabled = false;
                    serialPort1.PortName = comboBox1.Text;
                    serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                    serialPort1.DataBits = Convert.ToInt16(comboBox3.Text);

                    if (comboBox4.Text.Equals("None"))
                        serialPort1.Parity = System.IO.Ports.Parity.None;
                    else if (comboBox4.Text.Equals("Odd"))
                        serialPort1.Parity = System.IO.Ports.Parity.Odd;
                    else if (comboBox4.Text.Equals("Even"))
                        serialPort1.Parity = System.IO.Ports.Parity.Even;
                    else if (comboBox4.Text.Equals("Mark"))
                        serialPort1.Parity = System.IO.Ports.Parity.Mark;
                    else if (comboBox4.Text.Equals("Space"))
                        serialPort1.Parity = System.IO.Ports.Parity.Space;

                    if (comboBox5.Text.Equals("1"))
                        serialPort1.StopBits = System.IO.Ports.StopBits.One;
                    else if (comboBox5.Text.Equals("1.5"))
                        serialPort1.StopBits = System.IO.Ports.StopBits.OnePointFive;
                    else if (comboBox5.Text.Equals("2"))
                        serialPort1.StopBits = System.IO.Ports.StopBits.Two;

                    serialPort1.Open();     //打开串口
                    openport_bt.Text = "关闭串口";
                    openport_bt.BackColor = Color.Firebrick;
                }
            }
            catch (Exception ex)
            {
                //捕获可能发生的异常并进行处理

                //捕获到异常，创建一个新的对象，之前的不可以再用
                serialPort1 = new System.IO.Ports.SerialPort();
                //刷新COM口选项
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
                //响铃并显示异常给用户
                System.Media.SystemSounds.Beep.Play();
                openport_bt.Text = "打开串口";
                openport_bt.BackColor = Color.ForestGreen;
                MessageBox.Show(ex.Message);
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
                comboBox4.Enabled = true;
                comboBox5.Enabled = true;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {

                try
                {
                    //首先判断串口是否开启
                    if (serialPort1.IsOpen)
                    {
                        //串口处于开启状态，将发送区文本发送
                        //serialPort1.Write(textBox_send.Text);
                        if (!t.Enabled)
                        {
                            if (t.Interval < 500 || t.Interval > 10000)
                            {
                                MessageBox.Show("发送间隔不能小于500,也不能大于10000，太快也没用，AIDA64反正都是间隔一秒传数据", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                t.Interval = (double)textSendInterval.Value;
                                t.Start();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //捕获到异常，创建一个新的对象，之前的不可以再用
                    serialPort1 = new System.IO.Ports.SerialPort();
                    //刷新COM口选项
                    comboBox1.Items.Clear();
                    comboBox1.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
                    //响铃并显示异常给用户
                    System.Media.SystemSounds.Beep.Play();
                    openport_bt.Text = "打开串口";
                    openport_bt.BackColor = Color.ForestGreen;
                    MessageBox.Show(ex.Message);
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = true;
                    comboBox5.Enabled = true;
                }

        }

        private void SerialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            try
            {
                //因为要访问UI资源，所以需要使用invoke方式同步ui
                this.Invoke((EventHandler)(delegate
                {
                    textBox_receive.AppendText(serialPort1.ReadExisting());
                }
                   )
                );

            }
            catch (Exception ex)
            {
                //响铃并显示异常给用户
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show(ex.Message);

            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (t.Enabled)
            {
                t.Stop();
            }
        }

        private void textSendInterval_ValueChanged(object sender, EventArgs e)
        {
            if (textSendInterval.Value < 500 || textSendInterval.Value > 10000)
            {
                MessageBox.Show("发送间隔不能小于500,也不能大于10000，太快也没用，AIDA64反正都是间隔一秒传数据", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textSendInterval.Value = 500;
                return;
            }
            t.Stop();
            t.Interval = (double)textSendInterval.Value;
            t.Start();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.Write(textBox_send.Text);
            } else
            {
                MessageBox.Show("串口没打开");
            }
        }
    }

}
