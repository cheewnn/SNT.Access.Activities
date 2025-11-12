using SNT.Access.Activities.Helpers;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SNT.Access.Activities
{
    public class OpenAccessDB : CodeActivity
    {
        [RequiredArgument]
        public InArgument<string> AccdbPath { get; set; }

        public OutArgument<AccessSession> Session { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var accdbPath = AccdbPath.Get(context);
            var session = new SNT.Access.Activities.Helpers.AccessSession(accdbPath);
            Session.Set(context, session);
        }
    }
}
