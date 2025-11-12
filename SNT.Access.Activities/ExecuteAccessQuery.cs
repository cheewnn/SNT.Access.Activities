using SNT.Access.Activities.Helpers;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SNT.Access.Activities
{
    public class ExecuteAccessQuery : CodeActivity
    {
        [RequiredArgument]
        public InArgument<AccessSession> Session { get; set; }

        [RequiredArgument]
        public InArgument<string> SqlText { get; set; }

        public OutArgument<DataTable> DtOut { get; set; }

        public OutArgument<int> RowsAffected { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var session = Session.Get(context);
            if (session == null)
                throw new InvalidOperationException("Access session not initialized. Use OpenAccessDB first.");

            var sql = SqlText.Get(context)?.Trim();
            if (string.IsNullOrWhiteSpace(sql))
                throw new ArgumentNullException(nameof(SqlText));

            if (sql.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
            {
                var dt = session.Query(sql);
                DtOut.Set(context, dt);
                RowsAffected.Set(context, 0);
            }
            else
            {
                var rows = session.Execute(sql);
                RowsAffected.Set(context, rows);
                DtOut.Set(context, null);
            }
        }
    }
}
