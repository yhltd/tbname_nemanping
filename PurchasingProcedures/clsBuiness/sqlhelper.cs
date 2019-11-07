using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace clsBuiness
{
   public class sqlhelper
    {
        
 
        ///私有属性:数据库连接字符串
        ///Data Source=(Local)          服务器地址
        ///Initial Catalog=SimpleMESDB  数据库名称
        ///User ID=sa                   数据库用户名
        ///Password=admin123456         数据库密码
        private static string connectionString = ConfigurationManager.ConnectionStrings["Purchasing"].ConnectionString.ToString();
 
 
 
        /// <summary>
        /// sqlHelp 的摘要说明:数据库访问助手类
        /// sqlHelper是从DAAB中提出的一个类，在这里进行了简化，DAAB是微软Enterprise Library的一部分，该库包含了大量大型应用程序
        /// 开发需要使用的库类。
        /// </summary>
 
 
        public void SqlHelp()
        {
            //无参构造函数
        }
 
        public static SqlConnection conn;
 
        //打开数据库连接
        public static void OpenConn()
        {
            string SqlCon = connectionString;//数据库连接字符串
            conn = new SqlConnection(SqlCon);
            if (conn.State.ToString().ToLower() == "open")
            {
 
            }
            else
            {
                conn.Open();
            }
        }

        //关闭数据库连接
        public static void CloseConn()
        {
            if (conn.State.ToString().ToLower() == "open")
            {
                //连接打开时
                conn.Close();
                conn.Dispose();
            }
        }
 
 
        // 读取数据
        public static SqlDataReader GetDataReaderValue(string sql)
        {
            OpenConn();
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            
            return dr;
        }
 
 
        // 返回DataSet
        public DataSet GetDataSetValue(string sql, string tableName)
        {
            OpenConn();
            SqlDataAdapter da;
            DataSet ds = new DataSet();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(ds, tableName);
            CloseConn();
            return ds;
        }
 
        //  返回DataView
        public DataView GetDataViewValue(string sql)
        {
            OpenConn();
            SqlDataAdapter da;
            DataSet ds = new DataSet();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(ds, "temp");
            CloseConn();
            return ds.Tables[0].DefaultView;
        }
 
        // 返回DataTable
        public DataTable GetDataTableValue(string sql)
        {
            OpenConn();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(sql, conn);
            da.Fill(dt);
            CloseConn();
            return dt;
        }
 
        // 执行一个SQL操作:添加、删除、更新操作
        public void ExecuteNonQuery(string sql)
        {
            OpenConn();
            SqlCommand cmd;
            cmd = new SqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            CloseConn();
        }
 
        // 执行一个SQL操作:添加、删除、更新操作，返回受影响的行
        public int ExecuteNonQueryCount(string sql)
        {
            OpenConn();
            SqlCommand cmd;
            cmd = new SqlCommand(sql, conn);
            int value = cmd.ExecuteNonQuery();
            return value;
        }
 
        //执行一条返回第一条记录第一列的SqlCommand命令
        public object ExecuteScalar(string sql)
        {
            OpenConn();
            SqlCommand cmd;
            cmd = new SqlCommand(sql, conn);
            object value = cmd.ExecuteScalar();
            return value;
        }
 
        // 返回记录数
        public int SqlServerRecordCount(string sql)
        {
            OpenConn();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = sql;
            cmd.Connection = conn;
            SqlDataReader dr;
            dr = cmd.ExecuteReader();
            int RecordCount = 0;
            while (dr.Read())
            {
                RecordCount = RecordCount + 1;
            }
            CloseConn();
            return RecordCount;
        }
 
 
        ///<summary>
        ///一些常用的函数
        ///</summary>
 
        //判断是否为数字
        public static bool IsNumber(string a)
        {
            if (string.IsNullOrEmpty(a))
                return false;
            foreach (char c in a)
            {
                if (!char.IsDigit(c))
                    return false;
            }
            return true;
        }
 
        // 过滤非法字符
        public static string GetSafeValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;
            value = Regex.Replace(value, @";", string.Empty);
            value = Regex.Replace(value, @"'", string.Empty);
            value = Regex.Replace(value, @"&", string.Empty);
            value = Regex.Replace(value, @"%20", string.Empty);
            value = Regex.Replace(value, @"--", string.Empty);
            value = Regex.Replace(value, @"==", string.Empty);
            value = Regex.Replace(value, @"<", string.Empty);
            value = Regex.Replace(value, @">", string.Empty);
            value = Regex.Replace(value, @"%", string.Empty);
            return value;
        }

    }
}
