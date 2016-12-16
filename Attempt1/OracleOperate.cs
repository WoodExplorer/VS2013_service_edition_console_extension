using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OracleClient;

namespace Attempt1
{
    class OracleOperate : DatabaseHelper
    {
        public OracleOperate(string strConnection, string dataType)
            : base(strConnection, dataType)
        {

        }

        public DataSet SelectDataSet(string strSqlCommand)
        {
            return base.Query(strSqlCommand);
        }

        public OracleDataReader SelectDataReader(string strSqlCommand)
        {
            return (OracleDataReader)base.ExecuteReader(strSqlCommand, null);
        }

        public int Delete(string strSqlCommand)
        {
            return ExecuteSql(strSqlCommand);
        }

        public int Insert(string strSqlCommand)
        {
            return ExecuteSql(strSqlCommand);
        }

        public void Update(string strSqlCommand)
        {
            ExecuteSql(strSqlCommand);
        }

        public void CreateTable(string strSqlCommand)
        {
            ExecuteSql(strSqlCommand);
        }

        public void DropTable(string strSqlCommand)
        {
            ExecuteSql(strSqlCommand);
        }

        public object GetScalarCmdRes(string strSqlCommand)
        {
            return GetSingle(strSqlCommand);
        }
    }
}
