using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SNT.Access.Activities.Helpers
{
    [Serializable]
    public sealed class AccessSession : IDisposable
    {
        [NonSerialized] private Application _access;
        [NonSerialized] private Database _db;

        public string FilePath { get; }

        public AccessSession(string accdbPath)
        {
            FilePath = accdbPath ?? throw new ArgumentNullException(nameof(accdbPath));
            _access = new Application
            {
                Visible = false,
                AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
            };
            _access.OpenCurrentDatabase(accdbPath, false);
            _db = _access.CurrentDb();

            // refresh links once on open
            RefreshLinkedTables(_db);
        }

        public System.Data.DataTable Query(string sql)
        {
            Microsoft.Office.Interop.Access.Dao.Recordset rs = null;
            try
            {
                rs = _db.OpenRecordset(sql,
                    Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenSnapshot,
                    Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly);

                return DaoRecordsetToDataTable(rs);
            }
            finally
            {
                ReleaseCom(rs);
            }
        }

        public int Execute(string sql)
        {
            _db.Execute(sql, RecordsetOptionEnum.dbFailOnError);
            return _db.RecordsAffected;
        }

        public void RefreshLinks() => RefreshLinkedTables(_db);

        public void Dispose()
        {
            try { _db?.Close(); } catch { }
            try { _access?.Quit(AcQuitOption.acQuitSaveNone); } catch { }
            ReleaseCom(_db);
            ReleaseCom(_access);
            _db = null;
            _access = null;
        }

        private static void RefreshLinkedTables(Database db)
        {
            if (db == null) return;
            TableDefs defs = null;
            try
            {
                defs = db.TableDefs;
                foreach (TableDef t in defs)
                {
                    if (!string.IsNullOrWhiteSpace(t.Connect))
                    {
                        try { t.RefreshLink(); } catch { }
                    }
                    ReleaseCom(t);
                }
            }
            finally
            {
                ReleaseCom(defs);
            }
        }

        private static DataTable DaoRecordsetToDataTable(Recordset rs)
        {
            var dt = new DataTable();
            if (rs == null || (rs.EOF && rs.BOF)) return dt;

            var fields = rs.Fields;
            foreach (Field f in fields)
            {
                dt.Columns.Add(f.Name, typeof(object));
            }

            rs.MoveFirst();
            while (!rs.EOF)
            {
                var row = dt.NewRow();
                for (int i = 0; i < fields.Count; i++)
                {
                    row[i] = fields[i].Value ?? DBNull.Value;
                }
                dt.Rows.Add(row);
                rs.MoveNext();
            }
            return dt;
        }

        private static void ReleaseCom(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.FinalReleaseComObject(o);
            }
            catch { }
        }
    }
}
