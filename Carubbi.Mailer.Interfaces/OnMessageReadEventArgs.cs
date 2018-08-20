using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Carubbi.Mailer.DTOs
{
    public class OnMessageReadEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
    }
}
