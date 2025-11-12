using System.Activities.DesignViewModels;
using System.Diagnostics;

namespace SNT.Access.Activities.ViewModels
{
    public class AccessActivitiesViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> AccdbPath { get; set; }
        public DesignInArgument<string> SqlText { get; set; }

        public DesignOutArgument<System.Data.DataTable> DtOut { get; set; }
        public DesignOutArgument<int> RowsAffected { get; set; }

        public AccessActivitiesViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            /*
             * The base call will initialize the properties of the view model with the values from the xaml or with the default values from the activity
             */
            base.InitializeModel();

            PersistValuesChangedDuringInit(); // mandatory call only when you change the values of properties during initialization

            var orderIndex = 0;

            AccdbPath.DisplayName = Resources.Access_DB_Path_DisplayName;
            AccdbPath.Tooltip = Resources.Access_DB_Path_Tooltip;
            AccdbPath.IsRequired = true;
            AccdbPath.IsPrincipal = true;
            AccdbPath.OrderIndex = orderIndex++;

            SqlText.DisplayName = Resources.SqlText_DisplayName;
            SqlText.Tooltip = Resources.SqlText_Tooltip;
            SqlText.IsRequired = true;
            SqlText.IsPrincipal = true;
            SqlText.OrderIndex = orderIndex++;

            DtOut.DisplayName = Resources.DtOut_DisplayName;
            DtOut.Tooltip = Resources.DtOut_Tooltip;
            DtOut.OrderIndex = orderIndex++;

            RowsAffected.DisplayName = Resources.RowsAffected_DisplayName;
            RowsAffected.Tooltip = Resources.RowsAffected_Tooltip;
            RowsAffected.OrderIndex = orderIndex++;
        }
    }

    public class OpenAccessViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> AccdbPath { get; set; }
        public DesignOutArgument<SNT.Access.Activities.Helpers.AccessSession> Session { get; set; }
        public OpenAccessViewModel(IDesignServices services) : base(services)
        {
        }
        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();
            var orderIndex = 0;

            AccdbPath.DisplayName = Resources.Access_DB_Path_DisplayName;
            AccdbPath.Tooltip = Resources.Access_DB_Path_Tooltip;
            AccdbPath.IsRequired = true;
            AccdbPath.IsPrincipal = true;
            AccdbPath.OrderIndex = orderIndex++;

            Session.DisplayName = Resources.Session_DisplayName;
            Session.Tooltip= Resources.Session_Tooltip;
            Session.IsRequired = true;
            Session.IsPrincipal = true;
            Session.OrderIndex = orderIndex++;


        }
    }

    public class ExecuteAccessQueryViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<SNT.Access.Activities.Helpers.AccessSession> Session { get; set; }
        public DesignInArgument<string> SqlText { get; set; }
        public DesignOutArgument<System.Data.DataTable> DtOut { get; set; }
        public DesignOutArgument<int> RowsAffected { get; set; }

        public ExecuteAccessQueryViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();
            var orderIndex = 0;

            Session.DisplayName = Resources.Session_DisplayName;
            Session.Tooltip = Resources.Session_Tooltip;
            Session.IsRequired = true;
            Session.IsPrincipal = true;
            Session.OrderIndex = orderIndex++;

            SqlText.DisplayName = Resources.SqlText_DisplayName;
            SqlText.Tooltip = Resources.SqlText_Tooltip;
            SqlText.IsRequired = true;
            SqlText.IsPrincipal = true;
            SqlText.OrderIndex = orderIndex++;

            DtOut.DisplayName = Resources.DtOut_DisplayName;
            DtOut.Tooltip = Resources.DtOut_Tooltip;
            DtOut.OrderIndex = orderIndex++;

            RowsAffected.DisplayName = Resources.RowsAffected_DisplayName;
            RowsAffected.Tooltip = Resources.RowsAffected_Tooltip;
            RowsAffected.OrderIndex = orderIndex++;
        }
    }

    public class CloseAccessViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<SNT.Access.Activities.Helpers.AccessSession> Session { get; set; }

        public CloseAccessViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();
            var orderIndex = 0;

            Session.DisplayName = Resources.Session_DisplayName;
            Session.Tooltip = Resources.Session_Tooltip;
            Session.IsRequired = true;
            Session.IsPrincipal = true;
            Session.OrderIndex = orderIndex++;
        }
    }
}
