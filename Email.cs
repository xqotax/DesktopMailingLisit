using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesktopMailingLisit
{
    public class Email
    {
        public int Id { get; set; }

        public string? EmailString { get; set; }

        public bool Include { get; set; }
    }
}
