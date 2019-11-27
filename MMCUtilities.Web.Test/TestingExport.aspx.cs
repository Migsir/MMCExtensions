using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using MMCSirUtilities;

namespace MMCUtilities.Web.Test {
    public partial class TestingExport : System.Web.UI.Page {


        public class DataTesting {

            public int Id { get; set; }
            public string DataName { get; set; }

            public DateTime CreationDate { get; set; }

        }




        protected void Page_Load(object sender, EventArgs e) {
            List<DataTesting> lst = new List<DataTesting>();
            lst.Add(new DataTesting { Id = 1, DataName = "Jose"  , CreationDate = DateTime.Now });
            lst.Add(new DataTesting { Id = 2, DataName = "Miguel", CreationDate = DateTime.Now });
            lst.Add(new DataTesting { Id = 3, DataName = "Mena", CreationDate = DateTime.Now });
            lst.Add(new DataTesting { Id = 4, DataName = "Cervantes", CreationDate = DateTime.Now });
            lst.Add(new DataTesting { Id = 5, DataName = "Data Main", CreationDate = DateTime.Now });
            MMCExtentions.ExportListToExcel(lst, "Testing", false);
        }
    }
}