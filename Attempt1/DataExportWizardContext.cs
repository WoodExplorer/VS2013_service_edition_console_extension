using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Attempt1
{
    class DataExportWizardContext
    {
        private string dbProviderFacotoryName;
        private string dbConnectionString;
        private string tableName;
        private string sqlCommand;
        private DataTable columns;
        private string fileType;
        private string fileName;

        public string DbProviderFacotoryName
        {
            get
            {
                return this.dbProviderFacotoryName;
            }
            set
            {
                this.dbProviderFacotoryName = value;
            }
        }

        public string DbConnectionString
        {
            get
            {
                return this.dbConnectionString;
            }
            set
            {
                this.dbConnectionString = value;
            }
        }

        public string TableName
        {
            get
            {
                return this.tableName;
            }
            set
            {
                this.tableName = value;
            }
        }

        public string SQLCommand
        {
            get
            {
                return this.sqlCommand;
            }
            set
            {
                this.sqlCommand = value;
            }
        }

        public DataTable Columns
        {
            get
            {
                return this.columns;
            }
            set
            {
                this.columns = value;
            }
        }

        public string FileType
        {
            get
            {
                return this.fileType;
            }
            set
            {
                this.fileType = value;
            }
        }

        public string FileName
        {
            get
            {
                return this.fileName;
            }
            set
            {
                this.fileName = value;
            }
        }

        public string BuildSQL()
        {
            if (this.sqlCommand != null && this.sqlCommand.Length > 0)
            {
                return this.sqlCommand;
            }
            return "";
        }

        public Spire.DataExport.Collections.StringListCollection SelectedColumns
        {
            get
            {
                Spire.DataExport.Collections.StringListCollection columns
                    = new Spire.DataExport.Collections.StringListCollection();
                foreach (DataRow columnMeta in this.columns.Rows)
                {
                    columns.Add(columnMeta["ColumnName"]);
                }

                return columns;
            }
        }
    }
}
