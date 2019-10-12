using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace DocuWareComObject.Interface
{
    [Guid("FF055DE1-7B36-4CF8-8583-62D2B30E8D96")]
    interface IDocumentMenegmentConector
    {
        bool ConectServer();
        bool DisconectServer();
        string GetAllFileCabinetsXML();
        string GetAllFileInCabinetXML(string cabinet_id);
        string GetAllFileXML();
        string[] GetDocumentArray();
        string GetDocumentName(string patch);

        void AddSearchField(string fild_name, string fild_value);
        void AddSearchField(string fild_name, string fild_value_from,string fild_value_to);

        string SearchExXML();
    }
}
