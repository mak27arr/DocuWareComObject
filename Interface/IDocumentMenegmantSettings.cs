using System;
using System.Collections.Generic;
using System.Text;

namespace DocuWareComObject.Interface
{
    interface IDocumentMenegmantSettings
    {
        string ServerUrl { get; set; }
        bool UseWindowsAuthentication { get; set; }
        string UserName { get; set; }
        string Password { get; set; }
    }
}
