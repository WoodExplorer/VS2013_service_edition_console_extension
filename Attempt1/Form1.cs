using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data.OracleClient;
using System.Windows.Forms;
using System.Diagnostics;

namespace Attempt1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void work()
        {
            IniFileOperation iniFileOperation = new IniFileOperation(Environment.CurrentDirectory + @"\wsyj.ini");

            if (!iniFileOperation.ExistINIFile())
            {
                MessageBox.Show("配置文件wsyj.ini打开失败！");
                return;
            }
            string strDataBase = iniFileOperation.IniReadValue("dbParam", "Initial Catalog");
            string strUserName = iniFileOperation.IniReadValue("dbParam", "User ID");
            string strPassword = iniFileOperation.IniReadValue("dbParam", "Password");
            Global.oracleConnectionString = @"Data Source=" + strDataBase +
                @";User ID=" + strUserName + "; Password=" + strPassword;

            OracleConnection sqlConn = new OracleConnection(Global.oracleConnectionString);
            OracleCommand sqlCommBlkInfo = sqlConn.CreateCommand();
            OracleCommand sqlCommUserInfo = sqlConn.CreateCommand();
            OracleCommand sqlCommQuery = sqlConn.CreateCommand();
            OracleCommand sqlCommProcTempScore = sqlConn.CreateCommand();
            OracleCommand sqlCommTempScore = sqlConn.CreateCommand();
            OracleDataReader dataReaderBlkInfo, dataReaderQuery;
            OracleDataAdapter dataAdapter = new OracleDataAdapter();
            DataSet dataSetTmpScore = new DataSet();

            try
            {
                sqlConn.Open();

                sqlCommBlkInfo.CommandText = "select * from DZINFO";
                dataReaderBlkInfo = sqlCommBlkInfo.ExecuteReader();

                if (!dataReaderBlkInfo.HasRows)
                {
                    MessageBox.Show("无记录", "错误");
                    return;
                }
                else
                {
                    while (dataReaderBlkInfo.Read())
                    {
                        string sName = dataReaderBlkInfo["NAME"].ToString();

                        Trace.WriteLine(sName, "Info");
                    }
                }
            }
            catch (OracleException Oe)
            {
                throw Oe;
            }
            finally
            {
                sqlConn.Close();
            }
        }
    }
}
