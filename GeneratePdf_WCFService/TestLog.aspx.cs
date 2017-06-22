using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using log4net;

namespace GeneratePdf_WCFService
{
    public partial class TestLog : System.Web.UI.Page
    {

        private static ILog logger = LogManager.GetLogger(typeof(TestLog));
        protected void Page_Load(object sender, EventArgs e)
        {

        }
    }
}