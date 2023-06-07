
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using ADODB;    // TODO: You will need to do something with either references or a NUGET package.  
using System.IO;

namespace ST_14e8da3cceef411896eec36af25726b2 // TODO:  replace this with your given namespace
{
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
        #region vdbc#
        public const string SQLServerProd = "Driver={SQL Server};Server=;Database=;Trusted_Connection=True;";               // sql 16
        public object nRecordsAffected = Type.DefaultBinder;
        public object oParams = Type.Missing;
        public ADODB.Connection ADOConn = new ADODB.Connection();
        public ADODB.Recordset ADOrec = new ADODB.Recordset();
        public ADODB.Recordset rs = new ADODB.Recordset();
        public ADODB.Command ADOcom = new ADODB.Command();
        public object rc;
        public string query;

        public void open_ADOConn(string connection_String = SQLServerProd)
        {
            ADODB.Connection conn;
            conn = new ADODB.Connection();
            conn.ConnectionString = connection_String;
            conn.Open();
            ADOConn = conn;
        }

        public void open_ADORec(string Native_SQL)
        {
            ADODB.Recordset rec = new ADODB.Recordset();
            rec.LockType = LockTypeEnum.adLockReadOnly;
            rec.CursorType = CursorTypeEnum.adOpenKeyset; // adOpenKeyset
            rec.ActiveConnection = ADOConn;
            //  rec.Source           = Native_SQL;
            rec.Open(Native_SQL, ADOConn, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockReadOnly, -1);
            ADOrec = rec;
            rs = ADOConn.Execute(Native_SQL, out rc, (int)ADODB.CommandTypeEnum.adCmdText);
        }
        public void adoCommand(string command_text, string connection_String = SQLServerProd)
        {
            ADODB.Command com = new ADODB.Command();
            open_ADOConn(connection_String);
            // ADOcom = CreateObject("ADODB.Command") 'late binding
            com.ActiveConnection = ADOConn;
            com.CommandType = (CommandTypeEnum)1; // adcmdtext
            com.CommandText = command_text;
            com.Execute(out nRecordsAffected, ref oParams, (int)ADODB.ExecuteOptionEnum.adExecuteNoRecords);
            //  com.Execute();
            close_ADOconn();
        }

        public void adoCommandHoldrec(string command_text, string connection_String = SQLServerProd)
        {
            open_ADOConn(connection_String); // we set our connection string incase we arnt using postgres
                                             // ADOcom = CreateObject("ADODB.Command")
            ADOcom.ActiveConnection = ADOConn;
            ADOcom.CommandType = (CommandTypeEnum)1; // adcmdtext
            ADOcom.CommandText = command_text;
            ADOrec = ADOcom.Execute(out nRecordsAffected, ref oParams, (int)ADODB.ExecuteOptionEnum.adExecuteNoRecords);
        }

        //	public string adolookup(string Native_SQL, string connection_String = SQLServerProd)
        //	{
        //		string adolookup;
        //		open_ADOConn(connection_String); // database connection layer
        //		open_ADORec(Native_SQL);            // create a recordset for our data
        //		if (ADOrec.BOF && ADOrec.EOF)
        //			adolookup = "0";                         // return 0
        //		else
        //			// adolookup = ADOrec.;// return the first tuple of the first column
        //			adolookup = rs.Collect[0].ToString();// return the first tuple of the first column
        //		close_ADOrec();                                  // close the recordset
        //		close_ADOconn();                                 // close the connection
        //		return adolookup;
        //	}
        public bool this_ado_rec_is_empty(string Native_SQL, string connection_String = SQLServerProd)
        {
            bool this_ado_rec_is_empty;
            open_ADOConn(connection_String);
            open_ADORec(Native_SQL);

            if (ADOrec.EOF & ADOrec.BOF)
                this_ado_rec_is_empty = true;
            else
                this_ado_rec_is_empty = false;

            close_ADOrec();
            close_ADOconn();
            return this_ado_rec_is_empty;
        }

        public void close_ADOrec()
        {
            if (!(ADOrec == null))
            {
                ADOrec.Close();
                ADOrec = null;
            }
        }

        public void close_ADOconn()
        {
            if (!(ADOConn == null))
            {
                ADOConn.Close();
                ADOConn = null;
            }
        }
        #endregion
        public void Main()
		{
            try
            {
                string outputPath = @"C:\Output\output.csv";
                var csvContent = new System.Text.StringBuilder();
                open_ADOConn();
                open_ADORec
                (
                    @"
                    SELECT 
                        my_table.field_1
                    ,   my_table.field_2
                    ,   my_table.field_3
                    FROM my_table
                "
                );
                csvContent.Append("field 1 header name");
                csvContent.Append(",feild 2 header name");
                csvContent.Append(",field 3 header name");
                csvContent.AppendLine();
                if (!ADOrec.BOF) { ADOrec.MoveFirst(); }
                do
                {
                    csvContent.Append(ADOrec.Collect[0].ToString());  // ("field 1"));
                    csvContent.Append(ADOrec.Collect[1].ToString());  // ("field 2"));
                    csvContent.Append(ADOrec.Collect[2].ToString());  // ("field 3"));
                    ADOrec.MoveNext();
                }
                while (!ADOrec.EOF);
                close_ADOrec();
                close_ADOconn();

                File.WriteAllText(outputPath, csvContent.ToString());
                Dts.TaskResult = (int)ScriptResults.Success;
            }
            catch
            {
                Dts.TaskResult = (int)ScriptResults.Failure;
            }
		}
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
	}
}
