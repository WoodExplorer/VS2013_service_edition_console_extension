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
            work();
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

                string teacherTable = "TEACHERS";
                string scoreTable = "SCORE_1_1_1";

                sqlCommBlkInfo.CommandText = "select * from " + scoreTable + "," + teacherTable + " t5 where userid_5 = t5.userid";
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
                        string paperNo = dataReaderBlkInfo["PAPERNO"].ToString();
                        string userId5 = dataReaderBlkInfo["USERID_5"].ToString();
                        string trueName = dataReaderBlkInfo["TRUENAME"].ToString();
                        string score5 = dataReaderBlkInfo["SCOREOF_5"].ToString(); // 虽然库中SCOREOF_5可能不是字符串类型，但是，这里我们调用.ToString()就可以了。

                        Trace.WriteLine(paperNo + ", " + userId5 + ", " + trueName + ", " + score5, "Info");
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
