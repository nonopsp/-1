using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 串口助手1
{
    class Class1
    {
        MySqlConnection conn;
        MySqlCommand cmd;
        MySqlDataAdapter adapter;
        MySqlTransaction trans;

        string strConn = "Database='data';Data Source = 'localhost'; User Id = 'root'; Password='000000';charset='utf8';pooling=true";

        private void Dispose()//用完后销毁内存
        {

            if (conn != null) { conn.Dispose(); conn = null; }
            if (cmd != null) { cmd.Dispose(); cmd = null; }
            if (adapter != null) { adapter.Dispose(); adapter = null; }
            if ((trans != null)) { trans.Dispose(); trans = null; }

        }

        public DataTable Get_Table()//查看数据表
        {
            DataTable dt = new DataTable();
            string str = "show tables from data";
            try
            {
                conn = new MySqlConnection(strConn);
                conn.Open();

                adapter = new MySqlDataAdapter(str, conn);
                adapter.Fill(dt);

            }
            catch (Exception)
            {
                // Logger.Error("数据库不存在或其他问题");
            }
            finally
            {
                Dispose();
            }

            return dt;
        }
        public DataTable Get_Data(string date)//查看数据
        {
            lock (this)
            {
                DataTable dt = new DataTable();
                string str = $"select *from {date}";

                try
                {
                    conn = new MySqlConnection(strConn);
                    conn.Open();

                    adapter = new MySqlDataAdapter(str, conn);
                    adapter.Fill(dt);
                }
                catch (Exception)
                {
                    //Logger.Error("数据库不存在或其他问题");
                }
                finally
                {
                    Dispose();
                }

                return dt;
            }
        }
        public DataTable Get_Data(string date, int row)//查看数据
        {
            lock (this)
            {
                DataTable dt = new DataTable();
                string str = $"select*from {date} limit {row}";

                try
                {
                    conn = new MySqlConnection(strConn);
                    conn.Open();

                    adapter = new MySqlDataAdapter(str, conn);
                    adapter.Fill(dt);
                }
                catch (Exception)
                {
                    //Logger.Error("数据库不存在或其他问题");
                }
                finally
                {
                    Dispose();
                }

                return dt;
            }
        }
        public void Save_Data_alarm(string dt, double wave, string state)//保存数据
        {

            lock (this)
            {
                string time = DateTime.Now.ToString("_yyyy_MM_dd_HH") + "_alarm";
                string str = $"insert into {time} " +
                $"values('{DateTime.Now.ToString("HH:mm")}','{wave}','{state}')";

                try
                {
                    conn = new MySqlConnection(strConn);
                    conn.Open();
                    cmd = new MySqlCommand(str, conn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    //Logger.Error(ex.Message);
                    if (ex.Message == $"Table 'data.{time}' doesn't exist")
                    {
                        creatTable_alarm();
                        //Logger.Message("已创建表格");
                    }
                }
                finally
                {
                    Dispose();
                }
            }
        }

        public void Save_Data(string dt, double wave, string state)//保存数据
        {

            lock (this)
            {
                string time = DateTime.Now.ToString("_yyyy_MM_dd_HH");
                string str = $"insert into {DateTime.Now.ToString("_yyyy_MM_dd_HH")} " +
                $"values('{DateTime.Now.ToString("HH:mm")}','{wave}','{state}')";

                try
                {
                    conn = new MySqlConnection(strConn);
                    conn.Open();
                    cmd = new MySqlCommand(str, conn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    //Logger.Error(ex.Message);
                    if (ex.Message == $"Table 'data.{time}' doesn't exist")
                    {
                        creatTable();
                        //Logger.Message("已创建表格");
                    }
                }
                finally
                {
                    Dispose();
                }
            }
        }
        public void creatTable_alarm()//创建每日数据表
        {
            string time = DateTime.Now.ToString("_yyyy_MM_dd_HH") + "_alarm";
            string str = "CREATE TABLE " + time + " (time nvarchar(50) not null,wave nvarchar(100),state nvarchar(50))";//被创建数据表
            try
            {
                conn = new MySqlConnection(strConn);
                conn.Open();

                cmd = new MySqlCommand(str, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                //Logger.Error(ex.Message);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Dispose();
            }
        }
        public void creatTable()//创建每日数据表
        {
            string time = DateTime.Now.ToString("_yyyy_MM_dd_HH");
            string str = "CREATE TABLE " + time + " (time nvarchar(50) not null,wave nvarchar(100),state nvarchar(50))";//被创建数据表
            try
            {
                conn = new MySqlConnection(strConn);
                conn.Open();

                cmd = new MySqlCommand(str, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                //Logger.Error(ex.Message);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Dispose();
            }
        }
    }
}
