using System;
using System.Collections.Generic;
using System.Text;

namespace UdemyWordToPdf.Consumer
{
    class MessageWordToPdf
    {
        public byte[] WordByte { get; set; }
        public string Email { get; set; }
        public string FileName { get; set; }
    }
}
