using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MarsAddIn.HelperClasses
{
    internal static class DatabaseClass
    {
        private const int MaxExcelRows = 1048576;
        private const int MaxVarcharLength = 64 * 1024;

        public static List<DatabaseObject> GetDatabaseObjects(
          string serverName,
          string databaseName = "",
          string tableName = "",
          bool includeView = false)
        {
            SqlConnection conn = ConnectDb(serverName);
            string query;
            if (tableName != "")
            {
                string str1 = tableName;
                if (tableName.IndexOf(".") > 0)
                    str1 = tableName.Substring(tableName.IndexOf(".") + 1).Replace("[", "").Replace("]", "");
                string str2 = "dbo";
                if (tableName.IndexOf(".") > 0)
                    str2 = tableName.Substring(1, tableName.IndexOf(".") - 1).Replace("[", "").Replace("]", "");
                query = "SELECT COLUMN_NAME[ColumnName], DATA_TYPE[DataType], ISNULL(CHARACTER_MAXIMUM_LENGTH, 0) [Length], clmns.is_identity [IsIdentity] FROM[" + databaseName + "].[INFORMATION_SCHEMA].[COLUMNS] INNER JOIN [" + databaseName + "].[sys].[tables] tbl ON tbl.[name] = TABLE_NAME INNER JOIN [" + databaseName + "].[sys].[columns] clmns ON clmns.[object_id] = tbl.[object_id] AND clmns.[name] = COLUMN_NAME WHERE TABLE_NAME = '" + str1 + "' AND TABLE_SCHEMA = '" + str2 + "' ORDER BY ORDINAL_POSITION ASC;";
            }
            else if (databaseName != "")
            {
                AzObject.DatabaseName = databaseName;
                string str;
                if (!includeView)
                    str = "SELECT [TableName] FROM (SELECT  CONCAT('[', schm.[name], '].[',  tbl.[name], ']') [TableName], 0 [ObjOrder] FROM [" + databaseName + "].[sys].[tables] tbl INNER JOIN  [" + databaseName + "].[sys].[schemas] schm ON schm.[schema_id] = tbl.[schema_id] ) _q1 ORDER BY _q1.ObjOrder ASC, _q1.[TableName] ASC;";
                else
                    str = "SELECT [TableName] FROM (SELECT  CONCAT('[', schm.[name], '].[',  tbl.[name], ']') [TableName], 0 [ObjOrder] FROM [" + databaseName + "].[sys].[tables] tbl INNER JOIN  [" + databaseName + "].[sys].[schemas] schm ON schm.[schema_id] = tbl.[schema_id]  UNION ALL SELECT  CONCAT('[', schm.[name], '].[',  vws.[name], ']') [TableName], 1 [ObjOrder] FROM [" + databaseName + "].[sys].[views] vws INNER JOIN  [" + databaseName + "].[sys].[schemas] schm ON schm.[schema_id] = vws.[schema_id] ) _q1 ORDER BY _q1.ObjOrder ASC, _q1.[TableName] ASC;";
                query = str;
            }
            else
                query = "SELECT [name] [DatabaseName] FROM [sys].[databases] WHERE [name] NOT IN ('master', 'tempdb', 'model', 'msdb', 'SSISDB') ORDER BY [name] ASC;";
            return conn == null ? new List<DatabaseObject>() : DatabaseClass.QueryDatabase(conn, query);
        }

        private static SqlConnection ConnectDb(string serverName, string databaseName = "master")
        {
            string connectionString;

            if (serverName.EndsWith(".database.windows.net", StringComparison.OrdinalIgnoreCase))
            {
                AzObject.ServerName = serverName;
                connectionString =
                    $"Server=tcp:{serverName},1433; Initial Catalog={databaseName}; Persist Security Info=False; MultipleActiveResultSets=False; Encrypt=True; TrustServerCertificate=False; Connection Timeout=30; Authentication=Active Directory Interactive;";
            }
            else
                connectionString = $"Data Source={serverName};Initial Catalog={databaseName};Integrated Security=SSPI;Encrypt=True;TrustServerCertificate=True;";

            var cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
                return cnn;
            }
            catch (Exception ex)
            {
                cnn.Dispose();
                var message = ex.Message;
                var inner = ex.InnerException;
                while (inner != null)
                {
                    message += "\n\n" + inner.Message;
                    inner = inner.InnerException;
                }
                MessageBox.Show(message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private static List<DatabaseObject> QueryDatabase(SqlConnection conn, string query)
        {
            var dbList = new List<DatabaseObject>();

            using (var cmd = new SqlCommand { CommandText = query, CommandType = CommandType.Text, Connection = conn })
            {
                SqlDataReader reader;
                try
                {
                    reader = cmd.ExecuteReader();
                }
                catch (Exception)
                {
                    var retryConn = ConnectDb(AzObject.ServerName, AzObject.DatabaseName);
                    if (retryConn == null) return dbList;

                    using (var retryCmd = new SqlCommand { CommandText = query, CommandType = CommandType.Text, Connection = retryConn })
                    {
                        try
                        {
                            reader = retryCmd.ExecuteReader();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Database Query Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            retryConn.Close();
                            return dbList;
                        }

                        ReadDatabaseObjects(reader, query, dbList);
                        reader.Close();
                        retryConn.Close();
                        return dbList;
                    }
                }

                ReadDatabaseObjects(reader, query, dbList);
                reader.Close();
            }

            conn.Close();
            return dbList;
        }

        private static void ReadDatabaseObjects(SqlDataReader reader, string query, List<DatabaseObject> dbList)
        {
            while (reader.Read())
            {
                if (query.Contains("DatabaseName"))
                    dbList.Add(new DatabaseObject { DatabaseName = reader[0].ToString() });
                else if (query.Contains("TableName"))
                    dbList.Add(new DatabaseObject { TableName = reader[0].ToString() });
                else
                {
                    var dataType = reader[1].ToString();
                    var length = (int)reader[2];

                    if (dataType.Contains("varchar") && length == -1)
                        length = MaxVarcharLength;

                    dbList.Add(new DatabaseObject { ColumnName = reader[0].ToString(), DataType = dataType, ColumnLength = length, IsIdentity = (bool)reader[3] });
                }
            }
        }


        private static bool ExecuteOnDatabase(SqlConnection conn, string query)
        {
            try
            {
                using (var cmd = new SqlCommand { CommandText = query, CommandType = CommandType.Text, Connection = conn })
                {
                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Execute Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                conn.Close();
            }
        }


        public static string[] GetDatabases(List<DatabaseObject> dbobjs)
        {
            return dbobjs.Select(x => x.DatabaseName).ToArray();
        }

        public static string[] GetTables(List<DatabaseObject> dbobjs)
        {
            return dbobjs.Select(x => x.TableName).ToArray();
        }

        public static List<string> GetExcelColumns(List<string> headerList, bool withHeader = true)
        {
            if (headerList.Count > 0) return headerList;

            Excel.Range headerRng = null;
            try
            {
                Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
                var lastColumnInt = wks.Cells[1, 255].End(Excel.XlDirection.xlToLeft).Column();
                headerRng = wks.Range[wks.Cells[1, 1], wks.Cells[1, lastColumnInt]];
            }
            catch (Exception)
            {
                // ignored
            }

            if (headerRng == null) return headerList;

            var headerListL = new List<string>();

            foreach (Excel.Range rng in headerRng)

            {
                if (rng.Value == null || rng.Value.ToString().Length <= 0) continue;

                if (withHeader)
                    headerListL.Add(rng.Value.ToString());
                else
                {
                    var address = rng.Address;
                    headerListL.Add("Col " + address.Substring(1, address.LastIndexOf("$", StringComparison.Ordinal) - 1).Replace("$", ""));
                }
            }
            return headerListL;
        }


        public static void ReadData(object server, string query, Excel.Range startRng, bool customQuery = false)
        {
            var conn = ConnectDb(server.ToString());

            if (conn == null || query == null) return;

            SqlConnection activeConn = conn;
            using (var cmd = new SqlCommand { CommandText = query, CommandType = CommandType.Text, Connection = conn })
            {
                SqlDataReader reader;
                try
                {
                    reader = cmd.ExecuteReader();
                }
                catch (Exception)
                {
                    activeConn = ConnectDb(AzObject.ServerName, AzObject.DatabaseName);
                    if (activeConn == null) return;

                    using (var retryCmd = new SqlCommand { CommandText = query, CommandType = CommandType.Text, Connection = activeConn })
                    {
                        try
                        {
                            reader = retryCmd.ExecuteReader();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Database Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            activeConn.Close();
                            return;
                        }

                        WriteReaderToRange(reader, startRng, customQuery);
                        reader.Close();
                        activeConn.Close();
                        return;
                    }
                }

                WriteReaderToRange(reader, startRng, customQuery);
                reader.Close();
            }

            activeConn.Close();
        }

        private static void WriteReaderToRange(SqlDataReader reader, Excel.Range startRng, bool customQuery)
        {
            if (customQuery)
            {
                var columns = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToList();
                for (var j = 0; j < columns.Count; j++)
                    startRng.Offset[-1, j].Value = columns[j];
            }

            var i = 0;
            while (reader.Read() && i < MaxExcelRows)
            {
                for (var j = 0; j < reader.FieldCount; j++)
                {
                    try
                    {
                        startRng.Offset[i, j].Value = reader[j];
                    }
                    catch (Exception)
                    {
                        startRng.Offset[i, j].Value = "'" + reader[j];
                    }
                }

                i++;
            }
        }

        public static WriteResultsObject WriteData(object server, object database, object table, string query,
            Excel.Range sourceRange, List<DatabaseObject> targetColumns)
        {
            var sourceRangeArray = (Array)sourceRange.Value;

            var conn = ConnectDb(server.ToString(), database.ToString());
            var rslt = new WriteResultsObject
            {
                BadTypeColumns = new List<int>(),
                TruncColumns = new List<int>(),
                WriteRows = 0
            };

           

            if (query != "")
                ExecuteOnDatabase(conn, query);

            var dt = new DataTable();

            foreach (var targetColumn in targetColumns)
                dt.Columns.Add(new DataColumn(targetColumn.ColumnName));


            for (var i = 2; i <= sourceRangeArray.GetLength(0); i++)
            {
                var listValues = new List<object>();
                for (var j = 1; j <= targetColumns.Count; j++)
                {
                    var currentValue = sourceRangeArray.GetValue(i, j);
                    var maxColumnLength = targetColumns[j - 1].ColumnLength;
                    var dataType = targetColumns[j - 1].DataType;

                    if (currentValue != null && (maxColumnLength > 0 && currentValue.ToString().Length > 0))
                    {
                        var currentValueS = currentValue.ToString();
                        var valLength = currentValueS.Length;
                        if (valLength > maxColumnLength)
                        {
                            currentValue = currentValueS.Substring(0, maxColumnLength);
                            if (!rslt.TruncColumns.Contains(j - 1))
                                rslt.TruncColumns.Add(j - 1);
                        }
                    }

                    if (currentValue != null && !rslt.BadTypeColumns.Contains(j - 1))
                    {
                        var strVal = currentValue.ToString();
                        var isValid = true;

                        switch (dataType)
                        {
                            case "int":
                                isValid = int.TryParse(strVal, out _);
                                break;
                            case "bigint":
                                isValid = long.TryParse(strVal, out _);
                                break;
                            case "decimal":
                            case "numeric":
                            case "money":
                            case "smallmoney":
                                isValid = decimal.TryParse(strVal, out _);
                                break;
                            case "float":
                            case "real":
                                isValid = double.TryParse(strVal, out _);
                                break;
                            case "bit":
                                isValid = strVal == "0" || strVal == "1" || 
                                          bool.TryParse(strVal, out _);
                                break;
                            case "datetime":
                            case "datetime2":
                            case "date":
                            case "smalldatetime":
                                isValid = DateTime.TryParse(strVal, out _);
                                break;
                        }

                        if (!isValid)
                            rslt.BadTypeColumns.Add(j - 1);
                    }


                    listValues.Add(currentValue);
                }

                var row = dt.NewRow();
                // ReSharper disable once CoVariantArrayConversion
                row.ItemArray = listValues.ToArray();
                dt.Rows.Add(row);
            }

            var bc = new SqlBulkCopy(conn.ConnectionString, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers)
            {
                DestinationTableName = table.ToString(),
                BatchSize = dt.Rows.Count,
            };
    

            foreach (DataColumn dataColumn in dt.Columns)
                bc.ColumnMappings.Add(dataColumn.ColumnName, dataColumn.ColumnName);

            var rowCount = dt.Rows.Count;
            try
            {
                bc.WriteToServer(dt);
            }
            catch (Exception e)
            {
                rowCount = 0;
                MessageBox.Show($@"Error inserting data: {e.Message}", caption: @"Error Inserting");
            }
            finally
            {
                bc.Close();
                conn.Close();
            }

            rslt.WriteRows = rowCount;

            return rslt;
        }

        public static void CreateTable(TableObject tbl)
        {

            var conn = ConnectDb(tbl.ServerName, tbl.DatabaseName);
            var queryTableExists = $"SELECT [name] [TableName] FROM [{tbl.DatabaseName}].[sys].[tables] WHERE [name] = '{tbl.TableName}'";
            var dbojb = QueryDatabase(conn, queryTableExists);

            if (dbojb.Count > 0)
            {
                if (tbl.Recreate)
                {
                    var queryDropTable = $"TRUNCATE TABLE [{tbl.TableName}]; DROP TABLE [{tbl.TableName}];";
                    conn = ConnectDb(tbl.ServerName, tbl.DatabaseName);

                    if (!ExecuteOnDatabase(conn, queryDropTable))
                    {
                        MessageBox.Show(@"ERROR: Dropping existing table", caption: @"Info");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show(@"ERROR: Table already exists", caption: @"Duplicate Table");
                    return;
                }
            }

            var queryCreateTable = $"CREATE TABLE [{tbl.TableName}] (";
            for (var i = 0; i < tbl.ColumnName.Count; i++)
            {
                queryCreateTable += $"[{tbl.ColumnName[i]}]";
                queryCreateTable += $" {tbl.DataType[i]}";

                if (tbl.DataType[i] == "VARCHAR")
                    queryCreateTable += "(" + (tbl.ColumnSize[i] == 0 ? "MAX" : tbl.ColumnSize[i].ToString()) + ")";

                if (tbl.IsId[i])
                    queryCreateTable += " IDENTITY(1, 1)";

                if (tbl.IsPrimaryKey[i])
                    queryCreateTable += " PRIMARY KEY";

                if (!tbl.IsPrimaryKey[i] && !tbl.IsId[i])
                    if (tbl.Nullable[i])
                        queryCreateTable += " NULL";
                    else
                        queryCreateTable += " NOT NULL";


                queryCreateTable += ",";
            }

            queryCreateTable = queryCreateTable.Substring(0, queryCreateTable.Length - 1) + ");";

            conn = ConnectDb(tbl.ServerName, tbl.DatabaseName);
            MessageBox.Show(ExecuteOnDatabase(conn, queryCreateTable)
                ? @"Table successfully created"
                : @"Table was not created", caption: @"Info");
        }
    }

    public static class AzObject
    {
        public static string ServerName { get; set; }
        public static string DatabaseName { get; set; }
    }

    public class DatabaseObject
    {
        public string DatabaseName { get; set; }
        public string TableName { get; set; }
        public string ColumnName { get; set; }
        public string DataType { get; set; }
        public int ColumnLength { get; set; }
        public bool IsIdentity { get; set; }
    }

    public class TableObject
    {
        public TableObject()
        {
            ColumnName = new List<string>();
            DataType = new List<string>();
            ColumnSize = new List<int>();
            IsPrimaryKey = new List<bool>();
            IsId = new List<bool>();
            Nullable = new List<bool>();
        }

        public string ServerName { get; set; }
        public string DatabaseName { get; set; }
        public string TableName { get; set; }
        public bool Recreate { get; set; }
        public List<string> ColumnName { get; }
        public List<string> DataType { get; }
        public List<int> ColumnSize { get; }
        public List<bool> IsPrimaryKey { get; }
        public List<bool> IsId { get; }
        public List<bool> Nullable { get; }
    }

    public class WriteResultsObject
    {
        public WriteResultsObject()
        {
            TruncColumns = new List<int>();
            BadTypeColumns = new List<int>();
        }
        public int WriteRows { get; set; }
        public List<int> TruncColumns { get; set; }
        public List<int> BadTypeColumns { get; set; }

    }
}
