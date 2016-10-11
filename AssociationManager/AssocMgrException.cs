using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FileMetadataAssociationManager
{
    class AssocMgrException : Exception
    {
        public string Description { get; set; }
        public Exception Exception { get; set; }
        public WindowsErrorCodes ErrorCode { get; set; }
        public string DisplayString
        {
            get
            {
                if (Exception == null)
                    return Description;
                else
                    return Description + "\r\n" + Exception.Message;
            }
        }
    }
}
;