using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using SNT.Access.Activities.Helpers;
using System.Activities;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace SNT.Access.Activities
{
    public class QueryAccessDB : CodeActivity // This base class exposes an OutArgument named Result
    {
        [RequiredArgument]
        public InArgument<string> AccdbPath { get; set; }

        [RequiredArgument]
        public InArgument<string> SqlText { get; set; }

        public OutArgument<System.Data.DataTable> DtOut { get; set; }

        public OutArgument<int> RowsAffected { get; set; }
        /*
         * The returned value will be used to set the value of the Result argument
         */
        protected override void Execute(CodeActivityContext context)
        {
            context.GetExecutorRuntime().LogMessage(new UiPath.Robot.Activities.Api.LogMessage(){
                EventType = TraceEventType.Information,
                Message = "Executing query for Microsoft Access"
            });

            var accdbPath = AccdbPath.Get(context);
            var sqlText = SqlText.Get(context);
            if (!File.Exists(accdbPath))
            {
                throw new ArgumentNullException("Access database not found. Please ensure that it is a valid path.");
            }
            var (dtOut, rowsAffected) = ExecuteInternal(accdbPath, sqlText);
            DtOut.Set(context, dtOut);
            RowsAffected.Set(context, rowsAffected);
        }

        /// <summary>
        /// Executes the provided SQL against the Access .accdb using DAO (no Access UI needed).
        /// - Before running, refreshes linked tables (TableDefs with non-empty Connect).
        /// - If it's a SELECT (very basic detection), returns a DataTable; RowsAffected = 0.
        /// - Otherwise executes as non-query and returns RowsAffected from DAO.
        /// </summary>
        public (DataTable dtOut, int rowsAffected) ExecuteInternal(string accdbPath, string sqlText)
        {
            Application access = null;
            Database db = null;
            Recordset rs = null;

            try
            {
                // 1) Launch Access host (invisible) so its Dataverse provider/ISAM loads
                access = new Application();
                access.Visible = false;
                access.OpenCurrentDatabase(accdbPath, false);

                // 2) Get DAO entry point *inside* Access
                db = access.CurrentDb();

                // 3) Refresh linked tables (best-effort)
                RefreshLinkedTables(db);

                // 4) Run query
                if (LooksLikeSelect(sqlText))
                {
                    rs = db.OpenRecordset(sqlText, RecordsetTypeEnum.dbOpenSnapshot, RecordsetOptionEnum.dbReadOnly);
                    return (DaoRecordsetToDataTable(rs), 0);
                }
                else
                {
                    db.Execute(sqlText, RecordsetOptionEnum.dbFailOnError);
                    // Optional: re-sync/refresh again
                    RefreshLinkedTables(db);
                    return (null, db.RecordsAffected);
                }
            }
            finally
            {
                TryReleaseCom(rs);
                TryReleaseCom(db);

                if (access != null)
                {
                    try { if (access.CurrentProject != null) access.CloseCurrentDatabase(); } catch { }
                    try { access.Quit(AcQuitOption.acQuitSaveNone); } catch { }
                }
                TryReleaseCom(access);
            }
        }
        private static bool LooksLikeSelect(string sql)
        {
            // Strip leading whitespace & SQL line comments ("-- ...")
            // This is intentionally lightweight; if you need robustness, plug in a proper SQL tokenizer
            var trimmed = sql.TrimStart();

            // Skip "-- comment" lines at the very start
            while (trimmed.StartsWith("--", StringComparison.Ordinal))
            {
                var newline = trimmed.IndexOf('\n');
                if (newline < 0) return false; // only a comment, no SQL
                trimmed = trimmed[(newline + 1)..].TrimStart();
            }

            // Also skip /* ... */ at the very beginning if present
            if (trimmed.StartsWith("/*", StringComparison.Ordinal))
            {
                var end = trimmed.IndexOf("*/", StringComparison.Ordinal);
                if (end >= 0)
                {
                    trimmed = trimmed[(end + 2)..].TrimStart();
                }
            }

            return trimmed.StartsWith("SELECT", true, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Best-effort refresh of linked tables. If a TableDef has a non-empty Connect, call RefreshLink.
        /// Swallows per-table refresh exceptions so one bad link doesn't kill the whole operation.
        /// </summary>
        private static void RefreshLinkedTables(Microsoft.Office.Interop.Access.Dao.Database db)
        {
            if (db == null) return;

            Microsoft.Office.Interop.Access.Dao.TableDefs tableDefs = null;
            Microsoft.Office.Interop.Access.Dao.TableDef t = null;

            try
            {
                tableDefs = db.TableDefs;

                for (int i = 0; i < tableDefs.Count; i++)
                {
                    t = tableDefs[i];

                    // Skip system tables and local tables
                    bool isSystem = (t.Attributes & (int)Microsoft.Office.Interop.Access.Dao.TableDefAttributeEnum.dbSystemObject) != 0;
                    if (isSystem) continue;

                    var connect = t.Connect;
                    if (!string.IsNullOrWhiteSpace(connect))
                    {
                        try
                        {
                            t.RefreshLink();
                        }
                        catch
                        {
                            // Ignore individual failures (network hiccups, broken links, permissions)
                            // You may log these if you wish.
                        }
                    }

                    TryReleaseCom(t);
                    t = null;
                }
            }
            finally
            {
                TryReleaseCom(t);
                TryReleaseCom(tableDefs);
            }
        }

        /// <summary>
        /// Converts a DAO Recordset into a System.Data.DataTable.
        /// </summary>
        private static System.Data.DataTable DaoRecordsetToDataTable(Microsoft.Office.Interop.Access.Dao.Recordset rs)
        {
            var dt = new System.Data.DataTable();

            if (rs == null || rs.EOF && rs.BOF)
                return dt;

            // Build columns
            Microsoft.Office.Interop.Access.Dao.Fields fields = null;
            Microsoft.Office.Interop.Access.Dao.Field field = null;

            try
            {
                fields = rs.Fields;
                for (int i = 0; i < fields.Count; i++)
                {
                    field = fields[i];
                    var colName = SafeColumnName(field.Name, dt);
                    var colType = MapDaoDataTypeToClr(field.Type);
                    dt.Columns.Add(colName, colType);
                    TryReleaseCom(field);
                    field = null;
                }
            }
            finally
            {
                TryReleaseCom(field);
                TryReleaseCom(fields);
            }

            // Populate rows
            rs.MoveFirst();
            while (!rs.EOF)
            {
                var row = dt.NewRow();

                Microsoft.Office.Interop.Access.Dao.Fields rowFields = null;
                Microsoft.Office.Interop.Access.Dao.Field rowField = null;

                try
                {
                    rowFields = rs.Fields;
                    for (int i = 0; i < rowFields.Count; i++)
                    {
                        rowField = rowFields[i];
                        var value = MapDaoValue(rowField);
                        dt.Columns[i].ReadOnly = false; // ensure writable
                        row[i] = value ?? System.DBNull.Value;
                        TryReleaseCom(rowField);
                        rowField = null;
                    }
                }
                finally
                {
                    TryReleaseCom(rowField);
                    TryReleaseCom(rowFields);
                }

                dt.Rows.Add(row);
                rs.MoveNext();
            }

            return dt;
        }

        private static string SafeColumnName(string desired, System.Data.DataTable dt)
        {
            if (!dt.Columns.Contains(desired))
                return desired;

            // De-duplicate: Name, Name_1, Name_2, ...
            int i = 1;
            string candidate;
            do
            {
                candidate = $"{desired}_{i++}";
            } while (dt.Columns.Contains(candidate));
            return candidate;
        }

        private static Type MapDaoDataTypeToClr(short daoType)
        {
            // https://learn.microsoft.com/office/client-developer/access/desktop-database-reference/datatypeenum-enumeration-dao
            switch ((Microsoft.Office.Interop.Access.Dao.DataTypeEnum)daoType)
            {
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBoolean: return typeof(bool);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbByte: return typeof(byte);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger: return typeof(short);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbLong: return typeof(int); // Note: DAO "Long" is 32-bit
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbCurrency: return typeof(decimal);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbSingle: return typeof(float);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble: return typeof(double);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDate: return typeof(DateTime);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDecimal: return typeof(decimal);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText: return typeof(string);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbLongBinary:
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBinary:
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbVarBinary:
                    return typeof(byte[]);
                case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbMemo:
                    return typeof(string);
                default:
                    return typeof(object);
            }
        }

        private static object MapDaoValue(Microsoft.Office.Interop.Access.Dao.Field f)
        {
            // DAO Null → C# null
            try
            {
                var v = f.Value;

                // Binary types need special handling
                switch ((Microsoft.Office.Interop.Access.Dao.DataTypeEnum)f.Type)
                {
                    case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBinary:
                    case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbVarBinary:
                    case Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbLongBinary:
                        // DAO returns an object; attempt to cast to byte[] when possible
                        if (v == null) return null;
                        if (v is byte[] bytes) return bytes;
                        // Some providers return an OLE object wrapper; last resort: try to copy via Convert.ChangeType or leave as object
                        return v;
                    default:
                        return v;
                }
            }
            catch (COMException)
            {
                // Some providers throw when accessing Value of certain fields (e.g., OLE Object when empty)
                return null;
            }
        }

        private static void TryReleaseCom(object com)
        {
            try
            {
                if (com != null && Marshal.IsComObject(com))
                {
                    Marshal.FinalReleaseComObject(com);
                }
            }
            catch
            {
                // ignore cleanup errors
            }
        }
    }
}
