using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.Common;
using System.Data;
using System.Reflection;

using System.Windows.Forms; // for messageBox

namespace Attempt1
{
    class ExportAlgorithm
    {
        private string fileName;
        private DbCommand command;
        private Spire.DataExport.Collections.StringListCollection columns;
        private string strSqlComm;
        private string fileType;

        public ExportAlgorithm(string _strSql, string _fileName, string _fileType)
        {
            strSqlComm = _strSql;
            fileName = _fileName;
            fileType = _fileType;
        }

        private void ExportData_XLS()
        {
            Spire.DataExport.XLS.CellExport cellExport
                = new Spire.DataExport.XLS.CellExport();
            cellExport.ActionAfterExport = Spire.DataExport.Common.ActionType.None;
            cellExport.AutoFitColWidth = true;
            cellExport.DataFormats.CultureName = "zh-CN";
            cellExport.DataFormats.Currency = "#,###,##0.00";
            cellExport.DataFormats.DateTime = "yyyy-M-d H:mm";
            cellExport.DataFormats.Float = "#,###,##0.00";
            cellExport.DataFormats.Integer = "#,###,##0";
            cellExport.DataFormats.Time = "H:mm";
            cellExport.SheetOptions.AggregateFormat.Font.Name = "Arial";
            cellExport.SheetOptions.CustomDataFormat.Font.Name = "Arial";
            cellExport.SheetOptions.DefaultFont.Name = "Arial";
            cellExport.SheetOptions.FooterFormat.Font.Name = "Arial";
            cellExport.SheetOptions.HeaderFormat.Font.Name = "Arial";
            cellExport.SheetOptions.HyperlinkFormat.Font.Color = Spire.DataExport.XLS.CellColor.Blue;
            cellExport.SheetOptions.HyperlinkFormat.Font.Name = "Arial";
            cellExport.SheetOptions.HyperlinkFormat.Font.Underline = Spire.DataExport.XLS.XlsFontUnderline.Single;
            cellExport.SheetOptions.NoteFormat.Alignment.Horizontal = Spire.DataExport.XLS.HorizontalAlignment.Left;
            cellExport.SheetOptions.NoteFormat.Alignment.Vertical = Spire.DataExport.XLS.VerticalAlignment.Top;
            cellExport.SheetOptions.NoteFormat.Font.Bold = true;
            cellExport.SheetOptions.NoteFormat.Font.Name = "Tahoma";
            cellExport.SheetOptions.NoteFormat.Font.Size = 8F;
            cellExport.SheetOptions.TitlesFormat.Font.Bold = true;
            cellExport.SheetOptions.TitlesFormat.Font.Name = "Arial";
            cellExport.Columns = columns;
            cellExport.SQLCommand = command;

            cellExport.FileName = fileName;
            cellExport.SaveToFile();
        }

        public void Export()
        {
            try
            {
                DataExportWizardContext context = new DataExportWizardContext();
                context.DbProviderFacotoryName = "System.Data.OracleClient";
                DbProviderFactory factory
                    = DbProviderFactories.GetFactory(context.DbProviderFacotoryName);
                using (DbConnection conn = factory.CreateConnection())
                {
                    context.DbConnectionString = Global.oracleConnectionString;
                    conn.ConnectionString = Global.oracleConnectionString;
                    conn.Open();

                    context.SQLCommand = strSqlComm;
                    command = factory.CreateCommand();
                    command.Connection = conn;
                    command.CommandText = context.BuildSQL();
                    DbDataReader result = command.ExecuteReader();
                    DataTable schema = result.GetSchemaTable();
                    context.Columns = schema;
                    context.FileType = fileType;
                    context.FileName = fileName;
                    String specialTypeExportMethodName
                        = String.Format("ExportData_{0}", context.FileType);
                    System.Reflection.MethodInfo specialTypeExportMethod
                        = this.GetType().GetMethod(specialTypeExportMethodName, BindingFlags.Instance | BindingFlags.NonPublic);
                    columns = context.SelectedColumns;
                    if (specialTypeExportMethod != null)
                    {
                        specialTypeExportMethod.Invoke(this, new Object[] { });
                    }
                }
            }
            catch (Exception e)
            {
                String message
                    = String.Format("Failed to export data.[{0}]", e.Message);
                MessageBox.Show(message, "DataExport Wizard");
            }
        }
    }
}
