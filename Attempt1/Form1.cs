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
        private OracleOperate oracleOperate;
        private OracleConnection sqlConn;

        private string cur_topLevelTeam;
        private string cur_blkStr;

        #region “界面级”函数们
        public Form1()
        {
            InitializeComponent();
            ContextMenuStrip listboxMenu = new ContextMenuStrip();
            ToolStripMenuItem rightMenu = new ToolStripMenuItem("Copy");
            rightMenu.Click += new EventHandler(Copy_Click);
            listboxMenu.Items.AddRange(new ToolStripItem[] { rightMenu });
            listBox_log.ContextMenuStrip = listboxMenu;

            //
            //listBox_log.Items.Add("STARTED");

            // 以上为“界面级”操作
            //
            init();
            oracleOperate = new OracleOperate(Global.oracleConnectionString, "Oracle");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            loadCourses();
        }
        private void comboBoxCourse_SelectedIndexChanged(object sender, EventArgs e)
        {
            Trace.WriteLine(comboBoxCourse.SelectedValue);
            Trace.WriteLine(comboBoxCourse.Text);

            loadBlks_wrapper();
        }

        private void comboBoxBlks_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strBlk = comboBoxBlks.Text;

            cur_blkStr = strBlk;
        }

        private void button_export_Click(object sender, EventArgs e)
        {
            export_excel();
        }

        /// <summary>
        /// ListBox中的项目的右击菜单中的“copy”选项的处理函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Copy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(listBox_log.Items[listBox_log.SelectedIndex].ToString());
        }
        
        #endregion

        /// <summary>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>

        #region “业务级”函数们
        private void init()
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

            //
            sqlConn = new OracleConnection(Global.oracleConnectionString);
            try
            {
                sqlConn.Open();
            }
            catch
            {
                MessageBox.Show("数据库打开失败！", "错误");
                return;
            }
        }

        private void loadCourses()
        {
            string strSqlCmd = "";

            strSqlCmd = "select * from dzinfo where checkflag='1' order by toplevelteam";
            OracleDataReader dataReader = oracleOperate.SelectDataReader(strSqlCmd);

            if (!dataReader.HasRows)
            {
                return;
            }
            comboBoxCourse.Items.Clear();
            while (dataReader.Read())
            {
                comboBoxCourse.Items.Add(dataReader["courseno"].ToString() +
                    "(" + dataReader["name"].ToString() + ")");
            }
            dataReader.Close();
            dataReader.Dispose();

            comboBoxCourse.SelectedIndex = 0;
            //loadBlks_wrapper(); // 有了上面的“comboBoxCourse.SelectedIndex = 0;”，loadBlks_wrapper就会被调用啦~
        }
        private void loadBlks_wrapper()
        {
            string strCourse = comboBoxCourse.Text;
            int left_parenthesis = strCourse.IndexOf("(");
            //int right_parenthesis = strCourse.IndexOf(")");
            string strCourseNo = strCourse.Substring(0, left_parenthesis);

            loadBlks(strCourseNo, comboBoxBlks);
        }
        private void loadBlks(string courseNo, ComboBox cb) 
        {
            string strBlkList = "";
            
            
            //t//OracleCommand sqlCommTempTable = sqlConn.CreateCommand();
            //t//OracleCommand sqlCommDelCreate = sqlConn.CreateCommand();
            OracleCommand sqlCommQuery = sqlConn.CreateCommand();

            // 根据科目名称获取所在大组负责的题块号，用循环从BlkInfo表中读取，解析后写入blk_Temp中
            sqlCommQuery.CommandText = "select TopLevelTeam from DzInfo where courseno='" + courseNo + "'";
            OracleDataReader dataReaderQuery = sqlCommQuery.ExecuteReader();

            if (dataReaderQuery.HasRows)
            {
                while (dataReaderQuery.Read())
                {
                    cur_topLevelTeam = dataReaderQuery["ToplevelTeam"].ToString();   // 获取大组号
                }

                sqlCommQuery.CommandText = "select distinct blkname from userinfo " +
                    "where role=4 and toplevelteam=" + cur_topLevelTeam;
                dataReaderQuery = sqlCommQuery.ExecuteReader();
                while (dataReaderQuery.Read())
                {
                    strBlkList = dataReaderQuery["blkname"].ToString();
                }
                strBlkList = strBlkList.Substring(strBlkList.IndexOf("}") + 2);
                // 如果最后有一个','，则需要先删除
                //if (strBlk[strBlk.Length - 1] == ',')
                //{
                //    strBlk = strBlk.Substring(0, strBlk.Length - 1);
                //}

                string[] blk_list__array = strBlkList.Split(',');

                cb.Items.Clear();
                foreach (string item in blk_list__array) {
                    if (0 == item.Trim().Length)
                    {
                        Trace.WriteLine("Got an item trimmed to empty string. This might be a normal phenomenon. Going to skip it.");
                        continue;
                    }
                    cb.Items.Add(item);
                }
                cb.SelectedIndex = 0;
            }
        }

        private string lookUpScoreTable(string strBlk)
        {
            OracleCommand sqlCommQuery = sqlConn.CreateCommand();


            sqlCommQuery.CommandText = "select blkNo,checkrule,blktblname,scoretblname,revaluateRule, MAXLOWSCR, DEFINERSN from BlkInfo " +
                "where blkno=" + strBlk + " and questionno=" + cur_topLevelTeam;
            OracleDataReader dataReaderQuery = sqlCommQuery.ExecuteReader();

            if (!dataReaderQuery.HasRows)
            {
                MessageBox.Show("题块信息表中没有记录，无法进行评分轨迹导出！", "错误");
                return null;
            }

            String strScoreTblName = "";
            int cnt = 0;
            while (dataReaderQuery.Read())
            {
                //strBlkTblName = dataReaderBlkInfo["Blktblname"].ToString();
                strScoreTblName = dataReaderQuery["scoretblname"].ToString();
                Trace.WriteLine(strScoreTblName);
                cnt++;
            }
            Debug.Assert(1 == cnt);

            return strScoreTblName;
        }

        private void export_excel()
        {
            string strFileName = "tmp.xls";
            //string fileType = "";
            
            System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
            saveDlg.Filter = "excel files (*.xls)|*.xls";
            //saveDlg.Filter = "text files (*.txt)|*.txt|word files (*.doc)|*.doc|excel files (*.xls)|*.xls|dbf files (*.dbf)|*.dbf";
            saveDlg.RestoreDirectory = true;
            
            if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                strFileName = saveDlg.FileName.ToString();

                listBox_log.Items.Add("开始导出");

                // 准备生成sql命令
                string teacherTable = "TEACHERS";
                string scoreTable = lookUpScoreTable(cur_blkStr);// "SCORE_1_1_1";
                string scoreUpperBound = textBox_upperBound.Text;

                // 生成sql命令
                string column_names = "PAPERNO as 试卷号, USERID_5 as 给分小组长账号, TRUENAME as 给分小组长真实姓名, SCOREOF_5 as 给分小组长分数";
                string command = "select " + column_names + " from " + scoreTable + "," + teacherTable + " t5" +
                                 " where userid_5 = t5.userid " +
                                 " and CHECKUSERID = -1" +
                                 " and (SCOREOF_5 >= 0 and SCOREOF_5 <= " + scoreUpperBound + ")" +
                                 " and STEPFLAG = 100" +
                                 " and STORETYPE = 58" +
                                 " order by PAPERNO";

                listBox_log.Items.Add(command);

                // 导出excel文件
                export_excel_impl(strFileName, command);
                listBox_log.Items.Add("导出完成");
            }
        }

        private void export_excel_impl(string file_name, string cmd)
        {

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

                

                #region
                //sqlCommBlkInfo.CommandText = "select * from " + scoreTable + "," + teacherTable + " t5 where userid_5 = t5.userid";
                //dataReaderBlkInfo = sqlCommBlkInfo.ExecuteReader();

                //if (!dataReaderBlkInfo.HasRows)
                //{
                //    MessageBox.Show("无记录", "错误");
                //    return;
                //}
                //else
                //{
                //    while (dataReaderBlkInfo.Read())
                //    {
                //        string paperNo = dataReaderBlkInfo["PAPERNO"].ToString();
                //        string userId5 = dataReaderBlkInfo["USERID_5"].ToString();
                //        string trueName = dataReaderBlkInfo["TRUENAME"].ToString();
                //        string score5 = dataReaderBlkInfo["SCOREOF_5"].ToString(); // 虽然库中SCOREOF_5可能不是字符串类型，但是，这里我们调用.ToString()就可以了。

                //        Trace.WriteLine(paperNo + ", " + userId5 + ", " + trueName + ", " + score5, "Info");
                //    }
                //}
                #endregion

                ExportAlgorithm exportAlgorithm
                        = new ExportAlgorithm(cmd, file_name, "XLS");
                exportAlgorithm.Export();
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
        #endregion
    }
}
