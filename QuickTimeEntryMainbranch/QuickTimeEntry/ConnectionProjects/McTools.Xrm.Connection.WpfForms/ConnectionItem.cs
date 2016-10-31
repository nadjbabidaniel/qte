using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace McTools.Xrm.Connection.WpfForms
{
    public class ConnectionItem
    {
        public string Name { get; set; }
        public string Server { get; set; }
        public string Organization { get; set; }
        public ConnectionDetail ConnectionInfo { get; set; }
    }
}
