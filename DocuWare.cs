using System;
using System.Text;
using System.Runtime.InteropServices;
using DocuWareComObject.Interface;
using System.Reflection;
using DocuWare.Platform.ServerClient;
using DocuWare.Class;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Linq;
using DocuWareComObject.Class;

namespace DocuWareComObject
{
    public class DocuWare : IDocumentMenegmentConector, IDocumentMenegmantSettings
    {
        //[assembly: ComVisible(true)]
        //[assembly: AssemblyKeyFile("mak27arr.snk")]

        private static Uri server_url;
        private ServiceConnection conector;
        public string ServerUrl
        {
            get { return server_url.AbsoluteUri; }
            set {
                try
                {
                    server_url = new Uri(value);
                }catch(Exception ex)
                {
                    Logging.Log("Error set server url" + value,ex);
                }
            }
        }
        public bool UseWindowsAuthentication { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        public bool ConectServer()
        {
            if (server_url == null)
                return false;
            try
            {
                if (UseWindowsAuthentication)
                {
                    conector = ServiceConnection.CreateWithWindowsAuthentication(server_url, System.Net.CredentialCache.DefaultCredentials);
                }
                else
                {
                    conector = ServiceConnection.Create(server_url, UserName, Password);
                }
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log("Error conect", ex);
                return false;
            }
        }

        public bool DisconectServer()
        {
            if (conector != null)
            {
                try
                {
                    conector.Disconnect();
                    conector = null;
                    return true;
                }
                catch (Exception ex)
                {
                    Logging.Log("Error disconect", ex);
                }
            }
            return false;
        }

        public List<FileCabinet> GetAllFileCabinets()
        {
            Organization org = null;
            try
            {
                org = conector.Organizations[0];
            }
            catch (Exception ex)
            {
                Logging.Log("Error get organization", ex);
            }
            if (org != null)
            {
                List<FileCabinet> cabinets = org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet;
                return cabinets;
            }
            else
            {
                return null;
            }
        }

        public string GetAllFileCabinetsXML()
        {
            List<FileCabinet> cabinets = GetAllFileCabinets();
            Serializator<List<FileCabinet>> ser = new Serializator<List<FileCabinet>>();
            if (cabinets != null)
                return ser.SerializationXMLString(cabinets);
            else
                return "Error";
        }

        public List<Document> GetAllFileInCabinet(string cabinet_id)
        {
            if (conector == null)
            {
                Logging.Log("Error get file from cabinet: not conected");
                return null;
            }
            DocumentsQueryResult queryResult = conector.GetFromDocumentsForDocumentsQueryResultAsync(cabinet_id).Result;
            List<Document> documentlist = new List<Document>();
            foreach (var document in queryResult.Items)
            {
                documentlist.Add(document);
            }
            return documentlist;
        }

        public string GetAllFileInCabinetXML(string cabinet_id)
        {
            List<Document> document = GetAllFileInCabinet(cabinet_id);
            if (document == null)
            {
                return "";
            }
            else
            {
                Serializator<List<Document>> ser = new Serializator<List<Document>>();
                return ser.SerializationXMLString(document);
            }
        }

        public List<Document> GetAllFile()
        {
            List<FileCabinet> cabinets = GetAllFileCabinets();
            if (cabinets == null)
            {
                Logging.Log("Error get cabinet list in action GetAllFile()");
                return null;
            }
            ConcurrentBag<Document> alldocument = new ConcurrentBag<Document>();
            Parallel.ForEach(cabinets, cabinet => {
                List<Document> file_in_cabinet = GetAllFileInCabinet(cabinet.Id);
                foreach (var doc in file_in_cabinet)
                {
                    alldocument.Add(doc);
                }
            });
            return alldocument.ToList();
        }

        public string GetAllFileXML()
        {
            List<Document> documents = GetAllFile();
            Serializator<List<Document>> ser = new Serializator<List<Document>>();
            return ser.SerializationXMLString(documents);
        }

        public string DownloadDocumet(string cabinet_id, string file_id, string downloadPatch)
        {
            if (conector == null)
            {
                Logging.Log("Not conected to server cant download");
            }
            try
            {
                Document document = conector.GetFromDocumentForDocumentAsync(int.Parse(file_id), cabinet_id).Result;
                var downloadResponse = document.PostToFileDownloadRelationForStreamAsync(
                    new FileDownload()
                    {
                        TargetFileType = FileDownloadType.Auto
                    }).Result;
                string save_url = downloadPatch + downloadResponse.ContentHeaders.ContentDisposition.FileName;
                using (var fileStream = File.Create(save_url))
                {
                    downloadResponse.Content.CopyTo(fileStream);
                    fileStream.Flush();
                }
                return save_url;
            }
            catch (Exception ex)
            {
                Logging.Log("Error download file", ex);
                return "";
            }
        }

        public string[] GetDocumentArray()
        {
            throw new NotImplementedException();
        }

        public string GetDocumentName(string Patch)
        {
            throw new NotImplementedException();
        }

    }
}
