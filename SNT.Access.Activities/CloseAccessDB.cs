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
    public class CloseAccessDB : CodeActivity
    {
        [RequiredArgument]
        public InArgument<AccessSession> Session { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var session = Session.Get(context);
            session?.Dispose();
        }
    }
}
