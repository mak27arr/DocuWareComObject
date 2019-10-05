using DocuWare.Class;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace DocuWareComObject.Class
{
    public class Serializator<T>
    {
        public string SerializationXMLString(T data_to_serializate)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(T));
                using (StringWriter textWriter = new StringWriter())
                {
                    formatter.Serialize(textWriter, data_to_serializate);
                    return textWriter.ToString();
                }
            }
            catch (Exception ex)
            {
                Logging.Log("Error serializer " + typeof(T).ToString(), ex);
                return "";
            }
        }
    }
}
