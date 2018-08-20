using System;

namespace Carubbi.Mailer.DTOs
{
    public class OnMessageReadEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
    }
}
