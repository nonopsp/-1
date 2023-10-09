using DataService;
using MySql.Data.MySqlClient;
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
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 串口助手1
{
    public partial class Form1 : Form
    {
        String serialPortName;
        AutoSizeForm autoSizeForm = new AutoSizeForm();
        bool isResize = false;
      
        MySqlConnection connection = new MySqlConnection("server=localhost;user=root;password=88No.12138;");
        MySqlCommand command;
        string query;
        private const int MAX_DATA_POINTS = 100; // 展示数据的最大值
        private int dataCount = 0;
        public Form1()
        {
            InitializeComponent();
            autoSizeForm.GetAllInitInfo(this.Controls[0], this);
            isResize = true;
            chart1.ChartAreas[0].CursorX.Interval = 0;
            chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] ports = System.IO.Ports.SerialPort.GetPortNames();//获取电脑上可用串口号
            comboBox1.Items.AddRange(ports);//给comboBox1添加数据
            comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//如果里面有数据,显示第0个
            comboBox2.Text = "38400";/*默认波特率:115200*/
            comboBox3.Text = "1";/*默认停止位:1*/
            comboBox4.Text = "8";/*默认数据位:8*/
            comboBox5.Text = "无";/*默认奇偶校验位:无*/
            
        }

        private void creattable()
        {
            connection.Open();
            switch (comboBox6.Text)
            {
                //case "PZEM-017":
                //    query = "CREATE TABLE IF NOT EXISTS Users1 (id INT AUTO_INCREMENT PRIMARY KEY, Column1 VARCHAR(50),  Column2 VARCHAR(50))";
                //    cmd = new MySqlCommand(query, connection);
                //    cmd.ExecuteNonQuery();
                //    break;
                case "YC1002-P8":
                    // 创建数据库
                     command = new MySqlCommand("CREATE DATABASE IF NOT EXISTS mydatabase;", connection);
                    command.ExecuteNonQuery();

                    // 切换到 mydatabase 数据库
                    command = new MySqlCommand("USE mydatabase;", connection);
                    command.ExecuteNonQuery();
                    query = "CREATE TABLE IF NOT EXISTS YC1002 (id INT AUTO_INCREMENT PRIMARY KEY, Column1 VARCHAR(50),  Column2 VARCHAR(50), Column3 VARCHAR(50), Column4 VARCHAR(50), Column5 VARCHAR(50))";
                    command = new MySqlCommand(query, connection);
                    command.ExecuteNonQuery();
                    break;
                case "YC105S":
                    // 创建数据库
                    command = new MySqlCommand("CREATE DATABASE IF NOT EXISTS mydatabase;", connection);
                    command.ExecuteNonQuery();

                    // 切换到 mydatabase 数据库
                    command = new MySqlCommand("USE mydatabase;", connection);
                    command.ExecuteNonQuery();
                    query = "CREATE TABLE IF NOT EXISTS YC105S (id INT AUTO_INCREMENT PRIMARY KEY, Column1 VARCHAR(50),  Column2 VARCHAR(50), Column3 VARCHAR(50), Column4 VARCHAR(50), Column5 VARCHAR(50))";
                    command = new MySqlCommand(query, connection);
                    command.ExecuteNonQuery();
                    break;
                case "zspt1003w":
                    // 创建数据库
                    command = new MySqlCommand("CREATE DATABASE IF NOT EXISTS mydatabase;", connection);
                    command.ExecuteNonQuery();

                    // 切换到 mydatabase 数据库
                    command = new MySqlCommand("USE mydatabase;", connection);
                    command.ExecuteNonQuery();
                    query = "CREATE TABLE IF NOT EXISTS zspt1003w (id INT AUTO_INCREMENT PRIMARY KEY, Column1 VARCHAR(50),  Column2 VARCHAR(50), Column3 VARCHAR(50), Column4 VARCHAR(50), Column5 VARCHAR(50))";
                    command = new MySqlCommand(query, connection);
                    command.ExecuteNonQuery();
                    break;
                default:
                    break;
            }
            // 创建表


            // 关闭连接
            connection.Close();
        }

       
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0219)
            {//设备改变
                if (m.WParam.ToInt32() == 0x8004)
                {//usb串口拔出
                    timer1.Enabled = false;
                    string[] ports = System.IO.Ports.SerialPort.GetPortNames();//重新获取串口
                    comboBox1.Items.Clear();//清除comboBox里面的数据
                    comboBox1.Items.AddRange(ports);//给comboBox1添加数据
                    if (button1.Text == "关闭串口")
                    {//用户打开过串口
                        if (!serialPort1.IsOpen)
                        {//用户打开的串口被关闭:说明热插拔是用户打开的串口
                            button1.Text = "打开串口";
                            serialPort1.Dispose();//释放掉原先的串口资源
                            comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//显示获取的第一个串口号
                        }
                        else
                        {
                            comboBox1.Text = serialPortName;//显示用户打开的那个串口号
                        }
                    }
                    else
                    {//用户没有打开过串口
                        comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//显示获取的第一个串口号
                    }
                }
                else if (m.WParam.ToInt32() == 0x8000)
                {//usb串口连接上
                    string[] ports = System.IO.Ports.SerialPort.GetPortNames();//重新获取串口
                    comboBox1.Items.Clear();
                    comboBox1.Items.AddRange(ports);
                    if (button1.Text == "关闭串口")
                    {//用户打开过一个串口
                        comboBox1.Text = serialPortName;//显示用户打开的那个串口号
                    }
                    else
                    {
                        comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//显示获取的第一个串口号
                    }
                }
            }
            base.WndProc(ref m);
        }

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string txt = "";
            byte[] masterBytes = new byte[21];
            int numBytesRead = 0;
            while (numBytesRead < 21)
            {
                numBytesRead += serialPort1.Read(masterBytes, numBytesRead, 21 - numBytesRead);
                Thread.Sleep(1);
            }

            if (numBytesRead == 21)
            {
                txt = "Rx: ";
                for (int i = 0; i < masterBytes.Length; i++)
                    txt += masterBytes[i].ToString("X2");
                this.BeginInvoke(new Action(() =>
                    textBox1.AppendText(txt + "\n")));
                this.Invoke((EventHandler)(delegate
                {
                    string displayDate = BitConverter.ToString(masterBytes).Replace("-", "");//转10进制字符串
                    if (displayDate != " ")
                    {
                        displayChart(displayDate);//显示波形
                    }

                }));
            }


            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();

        }

        double X_Value1, X_Value2, X_Value3, X_Value4;
        private void displayChart(string displayDate)
        {
            int index = dataGridView1.Rows.Add();

            switch (comboBox6.Text)
            {
                //    case "PZEM-017":  
                // X_Value1 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(14, 4))) / 10.0;
                //chart1.Series[0].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value1);//
                //dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString();
                //dataGridView1.Rows[index].Cells[1].Value = X_Value1;
                //Sql();break;
                case "YC1002-P8":
                    X_Value1 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(6, 4))) / 10.0;
                    X_Value2 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(10, 4))) / 10.0;
                    X_Value3 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(14, 4))) / 10.0;
                    X_Value4 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(18, 4))) / 10.0;
                    //X_Value5 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(22, 4))) / 10.0;
                    //X_Value6 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(26, 4))) / 10.0;
                    //X_Value7 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(30, 4))) / 10.0;
                    //X_Value8 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(34, 4))) / 10.0;
                    dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString();

                    label14.Text = X_Value1.ToString() + "℃";
                    label13.Text = X_Value2.ToString() + "℃";
                    label12.Text = X_Value3.ToString() + "℃";
                    label11.Text = X_Value4.ToString() + "℃";

                    dataGridView1.Rows[index].Cells[1].Value = X_Value1;
                    dataGridView1.Rows[index].Cells[2].Value = X_Value2;
                    dataGridView1.Rows[index].Cells[3].Value = X_Value3;
                    dataGridView1.Rows[index].Cells[4].Value = X_Value4;
                    //dataGridView1.Rows[index].Cells[5].Value = X_Value5;
                    //dataGridView1.Rows[index].Cells[6].Value = X_Value6;
                    //dataGridView1.Rows[index].Cells[7].Value = X_Value7;
                    //dataGridView1.Rows[index].Cells[8].Value = X_Value8;
                    chart1.Series[0].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value1); 
                    chart1.Series[1].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value2); 
                    chart1.Series[2].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value3);  
                    chart1.Series[3].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value4);
                    dataCount++; // 数据计数器加1
                    if (dataCount > MAX_DATA_POINTS) // 如果数据计数器超过了最大值
                    {
                        chart1.Series[0].Points.RemoveAt(0); // 移除最旧的数据
                        chart1.Series[1].Points.RemoveAt(0);
                        chart1.Series[2].Points.RemoveAt(0);
                        chart1.Series[3].Points.RemoveAt(0);
                        dataCount--; // 数据计数器减1
                        chart1.Invalidate();

                    }

                    Sql(); break;
                case "YC105S":
                    X_Value1 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(6, 4))) / 10.0;
                    X_Value2 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(10, 4))) / 10.0;
                    X_Value3 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(14, 4))) / 10.0;
                    X_Value4 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(18, 4))) / 10.0;
                    dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString();

                    label14.Text = X_Value1.ToString() + "℃";
                    label13.Text = X_Value2.ToString() + "℃";
                    label12.Text = X_Value3.ToString() + "℃";
                    label11.Text = X_Value4.ToString() + "℃";

                    dataGridView1.Rows[index].Cells[1].Value = X_Value1;
                    dataGridView1.Rows[index].Cells[2].Value = X_Value2;
                    dataGridView1.Rows[index].Cells[3].Value = X_Value3;
                    dataGridView1.Rows[index].Cells[4].Value = X_Value4;
                    chart1.Series[0].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value1);
                    chart1.Series[1].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value2);
                    chart1.Series[2].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value3);
                    chart1.Series[3].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value4);
                    //dataCount++; // 数据计数器加1
                    //if (dataCount > MAX_DATA_POINTS) // 如果数据计数器超过了最大值
                    //{
                    //    chart1.Series[0].Points.RemoveAt(0); // 移除最旧的数据
                    //    chart1.Series[1].Points.RemoveAt(0);
                    //    chart1.Series[2].Points.RemoveAt(0);
                    //    chart1.Series[3].Points.RemoveAt(0);
                    //    dataCount--; // 数据计数器减1
                    //    chart1.Invalidate();

                    //}
                    Sql(); break;
             
                case "zspt1003w":
                    X_Value1 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(6, 4))) / 10.0;
                    X_Value2 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(10, 4))) / 10.0;
                    X_Value3 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(14, 4))) / 10.0;
                    X_Value4 = Convert.ToInt32(ConvertHexToSIntStr(displayDate.Substring(18, 4))) / 10.0;
                    dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString();

                    label14.Text = X_Value1.ToString() + "℃";
                    label13.Text = X_Value2.ToString() + "℃";
                    label12.Text = X_Value3.ToString() + "℃";
                    label11.Text = X_Value4.ToString() + "℃";

                    dataGridView1.Rows[index].Cells[1].Value = X_Value1;
                    dataGridView1.Rows[index].Cells[2].Value = X_Value2;
                    dataGridView1.Rows[index].Cells[3].Value = X_Value3;
                    dataGridView1.Rows[index].Cells[4].Value = X_Value4;
                    chart1.Series[0].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value1);
                    chart1.Series[1].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value2);
                    chart1.Series[2].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value3);
                    chart1.Series[3].Points.AddXY(DateTime.Now.ToString("H:mm:ss"), X_Value4);
                    //dataCount++; // 数据计数器加1
                    //if (dataCount > MAX_DATA_POINTS) // 如果数据计数器超过了最大值
                    //{
                    //    chart1.Series[0].Points.RemoveAt(0); // 移除最旧的数据
                    //    chart1.Series[1].Points.RemoveAt(0);
                    //    chart1.Series[2].Points.RemoveAt(0);
                    //    chart1.Series[3].Points.RemoveAt(0);
                    //    dataCount--; // 数据计数器减1
                    //    chart1.Invalidate();

                    //}
                    Sql(); break;





                default: break;
            }



        }

        private object ConvertHexToSIntStr(string hexstr)
        {
            if (hexstr.StartsWith("0x"))
            {
                hexstr = hexstr.Substring(2);
            }

            //如果不是有效的16进制字符串或者字符串长度大于16或者是空，均返回NULL

            if (!IsHexadecimal(hexstr) || hexstr.Length > 16 || string.IsNullOrEmpty(hexstr))
            {
                return null;
            }
            if (hexstr.Length > 8)
            {
                return Convert.ToInt64(hexstr, 16).ToString();
            }
            else if (hexstr.Length > 4)
            {
                return Convert.ToInt32(hexstr, 16).ToString();
            }
            else
            {
                return Convert.ToInt16(hexstr, 16).ToString();
            }
        }

        private bool IsHexadecimal(string hexstr)
        {
            const string PATTERN = @"[A-Fa-f0-9]+$";
            return System.Text.RegularExpressions.Regex.IsMatch(hexstr, PATTERN);
        }

        public static string byteToHexStr(byte[] bytes)
        {
            string returnStr = "";
            try
            {
                if (bytes != null)
                {
                    for (int i = 0; i < bytes.Length; i++)
                    {
                        returnStr += bytes[i].ToString("X2");
                        returnStr += " ";//两个16进制用空格隔开,方便看数据
                    }
                }
                return returnStr;
            }
            catch (Exception)
            {
                return returnStr;
            }
        }





        private static byte[] strToToHexByte(String hexString)
        {
            int i;
            hexString = hexString.Replace(" ", "");//清除空格
            if ((hexString.Length % 2) != 0)//奇数个
            {
                byte[] returnBytes = new byte[(hexString.Length + 1) / 2];
                try
                {
                    for (i = 0; i < (hexString.Length - 1) / 2; i++)
                    {
                        returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
                    }
                    returnBytes[returnBytes.Length - 1] = Convert.ToByte(hexString.Substring(hexString.Length - 1, 1).PadLeft(2, '0'), 16);
                }
                catch
                {
                    MessageBox.Show("含有非16进制字符", "提示");
                    return null;
                }
                return returnBytes;
            }
            else
            {
                byte[] returnBytes = new byte[(hexString.Length) / 2];
                try
                {
                    for (i = 0; i < returnBytes.Length; i++)
                    {
                        returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
                    }
                }
                catch
                {
                    MessageBox.Show("含有非16进制字符", "提示");
                    return null;
                }
                return returnBytes;
            }
        }





        private void timer1_Tick(object sender, EventArgs e)
        {
            send();
        }
        //int state = 0;

        private void send()
        {

            if (serialPort1.IsOpen != true)
            {
                MessageBox.Show("请打开串口");

            }

            else
            {

                try
                {

                   
                    //serialPort1.Write(HexSendData0, 0, HexSendData0.Length); 

                    switch (comboBox6.Text)
                    {
                        //case "PZEM-017":
                        //    byte[] HexSendData0 = new byte[] { 0x01, 0x04, 0x00, 0x00, 0x00, 0x08 };
                        //    byte[] arr = Utility.CalculateCrc(HexSendData0, 6);//自动生成crc
                        //    List<byte> list = new List<byte>(HexSendData0);
                        //    list.Add(arr[0]);
                        //    list.Add(arr[1]);
                        //    byte[] byteArray = list.ToArray();
                        //    serialPort1.Write(byteArray, 0, byteArray.Length); break;
                        case "YC1002-P8":
                            byte[] HexSendData1 = new byte[] { 0x00, 0x03, 0x00, 0x08, 0x00, 0x08 };
                            byte[] arr1 = Utility.CalculateCrc(HexSendData1, 6);//自动生成crc
                            List<byte> list1 = new List<byte>(HexSendData1);
                            list1.Add(arr1[0]);
                            list1.Add(arr1[1]);
                            byte[] byteArray1 = list1.ToArray();
                            serialPort1.Write(byteArray1, 0, byteArray1.Length); break;
                        case "YC105S":
                            byte[] HexSendData2 = new byte[] { 0x02, 0x03, 0x00, 0x20, 0x00, 0x08 };
                            byte[] arr2 = Utility.CalculateCrc(HexSendData2, 6);//自动生成crc
                            List<byte> list2 = new List<byte>(HexSendData2);
                            list2.Add(arr2[0]);
                            list2.Add(arr2[1]);
                            byte[] byteArray2 = list2.ToArray();
                            serialPort1.Write(byteArray2, 0, byteArray2.Length); break;
                        case "zspt1003w":
                            byte[] HexSendData3 = new byte[] { 0x01, 0x03, 0x00, 0x00, 0x00, 0x08 };
                            byte[] arr3 = Utility.CalculateCrc(HexSendData3, 6);//自动生成crc
                            List<byte> list3 = new List<byte>(HexSendData3);
                            list3.Add(arr3[0]);
                            list3.Add(arr3[1]);
                            byte[] byteArray3 = list3.ToArray();
                            serialPort1.Write(byteArray3, 0, byteArray3.Length); break;
                        default: /*MessageBox.Show("请选择型号");*/
                            timer1.Enabled = false; MessageBox.Show("请选择型号,请重新打开串口"); break;
                    }

                }
                catch (Exception)
                {
                    MessageBox.Show("发送数据时发生错误！", "错误提示");
                    return;
                }
            }
        }

        private string ToHexString(byte[] bytes)
        {
            string hexString = string.Empty;

            if (bytes != null)

            {

                StringBuilder strB = new StringBuilder();

                for (int i = 0; i < bytes.Length; i++)

                {

                    strB.Append(bytes[i].ToString("X2"));

                }

                hexString = strB.ToString();

            }
            return hexString;
        }

        //private void button5_Click(object sender, EventArgs e)
        //{
        //    this.Invoke((EventHandler)(delegate
        //    {

        //        Sql();

        //    }));
        //    MessageBox.Show("导入成功");
        //}




        private void Sql()
        {



            // 打开连接
            connection.Open();


            switch (comboBox6.Text)
            {
                case "YC1002-P8":


                    query = "INSERT INTO YC1002 (Column1, Column2, Column3, Column4, Column5) VALUES (@Column1, @Column2, @Column3, @Column4, @Column5)";
                    command = new MySqlCommand(query, connection);
                    command.Parameters.AddWithValue("@Column1", DateTime.Now.ToString("H:mm:ss"));
                    command.Parameters.AddWithValue("@Column2", X_Value1);
                    command.Parameters.AddWithValue("@Column3", X_Value2);
                    command.Parameters.AddWithValue("@Column4", X_Value3);
                    command.Parameters.AddWithValue("@Column5", X_Value4);
                    //cmd.Parameters.AddWithValue("@Column6", X_Value5);
                    //cmd.Parameters.AddWithValue("@Column7", X_Value6);
                    //cmd.Parameters.AddWithValue("@Column8", X_Value7);
                    //cmd.Parameters.AddWithValue("@Column8", X_Value8);
                    command.ExecuteNonQuery();
                    break;
                case "YC105S":


                    query = "INSERT INTO YC105S (Column1, Column2, Column3, Column4, Column5) VALUES (@Column1, @Column2, @Column3, @Column4, @Column5)";
                    command = new MySqlCommand(query, connection);
                    command.Parameters.AddWithValue("@Column1", DateTime.Now.ToString("H:mm:ss"));
                    command.Parameters.AddWithValue("@Column2", X_Value1);
                    command.Parameters.AddWithValue("@Column3", X_Value2);
                    command.Parameters.AddWithValue("@Column4", X_Value3);
                    command.Parameters.AddWithValue("@Column5", X_Value4);
                    command.ExecuteNonQuery();
                    break;
                case "zspt1003w":
                    query = "INSERT INTO zspt1003w (Column1, Column2, Column3, Column4, Column5) VALUES (@Column1, @Column2, @Column3, @Column4, @Column5)";
                    command = new MySqlCommand(query, connection);
                    command.Parameters.AddWithValue("@Column1", DateTime.Now.ToString("H:mm:ss"));
                    command.Parameters.AddWithValue("@Column2", X_Value1);
                    command.Parameters.AddWithValue("@Column3", X_Value2);
                    command.Parameters.AddWithValue("@Column4", X_Value3);
                    command.Parameters.AddWithValue("@Column5", X_Value4);
                    command.ExecuteNonQuery();
                    break;
                default:
                    break;
            }
            // 创建表


            // 关闭连接
            connection.Close();
        }









        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (isResize)
            {
                autoSizeForm.ControlsChaneInit(this.Controls[0]);
                autoSizeForm.ControlsChange(this.Controls[0]);
            }

        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            //tabControl1.Width = this.ClientSize.Width - tabControl1.Left - 10; tabControl1.Height = this.ClientSize.Height - tabControl1.Top - 10;
            //foreach (TabPage tabPage in tabControl1.TabPages) { tabPage.Width = tabControl1.Width - tabPage.Left - 5; tabPage.Height = tabControl1.Height - tabPage.Top - 30; }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            if (button1.Text == "打开串口")
            {//如果按钮显示的是打开串口
                try
                {//防止意外错误
                    serialPort1.PortName = comboBox1.Text;//获取comboBox1要打开的串口号
                    serialPortName = comboBox1.Text;
                    serialPort1.BaudRate = int.Parse(comboBox2.Text);//获取comboBox2选择的波特率
                    serialPort1.DataBits = int.Parse(comboBox4.Text);//设置数据位
                    /*设置停止位*/
                    if (comboBox3.Text == "1") { serialPort1.StopBits = StopBits.One; }
                    else if (comboBox3.Text == "1.5") { serialPort1.StopBits = StopBits.OnePointFive; }
                    else if (comboBox3.Text == "2") { serialPort1.StopBits = StopBits.Two; }
                    /*设置奇偶校验*/
                    if (comboBox5.Text == "无") { serialPort1.Parity = Parity.None; }
                    else if (comboBox5.Text == "奇校验") { serialPort1.Parity = Parity.Odd; }
                    else if (comboBox5.Text == "偶校验") { serialPort1.Parity = Parity.Even; }

                    serialPort1.Open();//打开串口
                    button1.Text = "关闭串口";//按钮显示关闭串口
                    button6.Text = "关闭串口";//按钮显示关闭串口
                    creattable();
                    timer1.Enabled = true;
                    
                }
                catch (Exception err)
                {
                    MessageBox.Show("打开失败" + err.ToString(), "提示!");//对话框显示打开失败
                }
            }
            else
            {//要关闭串口
                try
                {//防止意外错误
                    serialPort1.Close();//关闭串口
                    timer1.Enabled = false;
                }
                catch (Exception) { }
                button1.Text = "打开串口";//按钮显示打开
                button6.Text = "打开串口";//按钮显示打开
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExportExcel();
        }

        private void ExportExcel()
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                if (!Directory.Exists(@"C:\BMDT"))
                    Directory.CreateDirectory(@"C:\BMDT");


                string fileName = "";
                string saveFileName = "";
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.DefaultExt = "xlsx";
                saveDialog.InitialDirectory = @"C:\BMDT";
                saveDialog.Filter = "Excel文件|*.xlsx";
                // saveDialog.FileName = fileName;
                saveDialog.FileName = "BMDT_Data_" + DateTime.Now.ToLongDateString().ToString();
                saveDialog.ShowDialog();
                saveFileName = saveDialog.FileName;



                if (saveFileName.IndexOf(":") < 0)
                {
                    this.Cursor = Cursors.Default;
                    return; //被点了取消
                }
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("无法创建Excel对象，您的电脑可能未安装Excel");
                    return;
                }
                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1 
                Microsoft.Office.Interop.Excel.Range range = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[8, 1]];

                //写入标题             
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                { worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText; }
                //写入数值
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        worksheet.Cells[r + 2, i + 1] = dataGridView1.Rows[r].Cells[i].Value;

                        if (this.dataGridView1.Rows[r].Cells[i].Style.BackColor == Color.Red)
                        {

                            range = worksheet.Range[worksheet.Cells[r + 2, i + 1], worksheet.Cells[r + 2, i + 1]];
                            range.Interior.ColorIndex = 10;

                        }



                    }
                    System.Windows.Forms.Application.DoEvents();
                }
                worksheet.Columns.EntireColumn.AutoFit();//列宽自适应

                MessageBox.Show(fileName + "资料保存成功", "提示", MessageBoxButtons.OK);
                if (saveFileName != "")
                {
                    try
                    {
                        workbook.Saved = true;
                        workbook.SaveCopyAs(saveFileName);  //fileSaved = true;  

                    }
                    catch (Exception ex)
                    {//fileSaved = false;                      
                        MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                    }
                }
                xlApp.Quit();
                GC.Collect();//强行销毁           

                this.Cursor = Cursors.Default;
            }
            catch (Exception e)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(e.ToString());

            }
        }

        private void cw1_CheckedChanged(object sender, EventArgs e)
        {
            if (cw1.Checked)
            {
                chart1.Series[0].Enabled = true;


            }
            else
            {
                chart1.Series[0].Enabled = false;


            }
        }

        private void cw2_CheckedChanged(object sender, EventArgs e)
        {
            if (cw2.Checked)
            {
                chart1.Series[1].Enabled = true;


            }
            else
            {
                chart1.Series[1].Enabled = false;


            }
        }

        private void cw3_CheckedChanged(object sender, EventArgs e)
        {
            if (cw3.Checked)
            {
                chart1.Series[2].Enabled = true;


            }
            else
            {
                chart1.Series[2].Enabled = false;


            }
        }

        private void cw4_CheckedChanged(object sender, EventArgs e)
        {
            if (cw4.Checked)
            {
                chart1.Series[3].Enabled = true;


            }
            else
            {
                chart1.Series[3].Enabled = false;


            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (button6.Text == "打开串口")
            {//如果按钮显示的是打开串口
                try
                {//防止意外错误
                    serialPort1.PortName = comboBox1.Text;//获取comboBox1要打开的串口号
                    serialPortName = comboBox1.Text;
                    serialPort1.BaudRate = int.Parse(comboBox2.Text);//获取comboBox2选择的波特率
                    serialPort1.DataBits = int.Parse(comboBox4.Text);//设置数据位
                    /*设置停止位*/
                    if (comboBox3.Text == "1") { serialPort1.StopBits = StopBits.One; }
                    else if (comboBox3.Text == "1.5") { serialPort1.StopBits = StopBits.OnePointFive; }
                    else if (comboBox3.Text == "2") { serialPort1.StopBits = StopBits.Two; }
                    /*设置奇偶校验*/
                    if (comboBox5.Text == "无") { serialPort1.Parity = Parity.None; }
                    else if (comboBox5.Text == "奇校验") { serialPort1.Parity = Parity.Odd; }
                    else if (comboBox5.Text == "偶校验") { serialPort1.Parity = Parity.Even; }

                    serialPort1.Open();//打开串口
                    button1.Text = "关闭串口";//按钮显示关闭串口
                    button6.Text = "关闭串口";//按钮显示关闭串口
                    creattable();
                    timer1.Enabled = true;
                }
                catch (Exception err)
                {
                    MessageBox.Show("打开失败" + err.ToString(), "提示!");//对话框显示打开失败
                }
            }
            else
            {//要关闭串口
                try
                {//防止意外错误
                    serialPort1.Close();//关闭串口
                    timer1.Enabled = false;
                }
                catch (Exception) { }
                button1.Text = "打开串口";//按钮显示打开
                button6.Text = "打开串口";//按钮显示打开
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            timer1.Interval = Convert.ToInt32(textBox3.Text);
        }

        private void chart1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset(0);
            }
        }

        private void button2_Click_3(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void button3_Click_3(object sender, EventArgs e)
        {
            String Str = textBox2.Text.ToString();//获取发送文本框里面的数据
            try
            {
                if (Str.Length > 0)
                {
                    if (checkBox2.Checked)//16进制发送
                    {
                        byte[] byt = strToToHexByte(Str);
                        serialPort1.Write(byt, 0, byt.Length);
                    }
                    else
                    {
                        serialPort1.Write(Str);//串口发送数据
                    }
                }
            }
            catch (Exception) { }
        }

        private void button4_Click_3(object sender, EventArgs e)
        {
            textBox2.Clear();
        }
    }
}
