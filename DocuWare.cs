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
using DocuWare.Services.Http;
using DocuWare.Platform.ServerClient;

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
                    server_url = new Uri("http://localhost");
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
            //Organization org = null;
            List<FileCabinet> cabinets = new List<FileCabinet>();

            foreach (Organization org in conector.Organizations)
            {
                try
                {
                    cabinets = cabinets.Concat(org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet).ToList();
                }catch(Exception ex)
                {
                    Logging.Log("Error get cabinets",ex);
                }
            }
            //try
            //{
            //    org = conector.Organizations[0];
            //}
            //catch (Exception ex)
            //{
            //    Logging.Log("Error get organization", ex);
            //}
            //if (org != null)
            //{
            //    List<FileCabinet> cabinets = org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet;
            //    return cabinets;
            //}
            //else
            //{
            //    return null;
            //}
            return cabinets;
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

        public List<Document> GetAllFileInCabinetBasked(string cabinet_id)
        {

            try
            {
                if (conector == null)
                {
                    Logging.Log("Error get file from cabinet: not conected");
                    return null;
                }
                DocumentsQueryResult queryResult = conector.GetFromDocumentsForDocumentsQueryResultAsync(cabinet_id, count:int.MaxValue).Result;
                List<Document> documentlist = new List<Document>();
                foreach (var document in queryResult.Items)
                {
                    documentlist.Add(document);
                }
                return documentlist;
            }
            catch (Exception ex)
            {
                Logging.Log("Error get file in cabinet", ex);
                return new List<Document>();
            }

        }

        public List<Document> GetAllFileInCabinet(string cabinet_id)
        {
            try
            {
                if (conector == null)
                {
                    Logging.Log("Error get file from cabinet: not conected");
                    return null;
                }
                DocumentsQueryResult queryResult = conector.GetFromDocumentsForDocumentsQueryResultAsync(cabinet_id, count: int.MaxValue).Result;
                List<Document> documentlist = new List<Document>();
                foreach (var document in queryResult.Items)
                {
                    documentlist.Add(document);
                }
                return documentlist;
            }catch(Exception ex)
            {
                Logging.Log("Error get file in cabinet",ex);
                return new List<Document>();
            }
        }

        public string GetAllFileInCabinetXML(string cabinet_id)
        {
            try
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
            }catch(Exception ex)
            {
                Logging.Log("Error get file in cabinet",ex);
                return "";
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
            Parallel.ForEach(cabinets, cabinet =>
            {
                try
                {
                    DocuWare docuWare = new DocuWare();
                    docuWare.ServerUrl = ServerUrl;
                    docuWare.UseWindowsAuthentication = this.UseWindowsAuthentication;
                    docuWare.UserName = UserName;
                    docuWare.Password = Password;
                    docuWare.ConectServer();
                    List<Document> file_in_cabinet = docuWare.GetAllFileInCabinet(cabinet.Id);
                    foreach (var doc in file_in_cabinet)
                    {
                        alldocument.Add(doc);
                    }
                    docuWare.DisconectServer();
                }
                catch (Exception ex)
                {
                    Logging.Log("Error thread stop add", ex);
                }
            });

            //List<Document> alldocument = new List<Document>();
            //foreach (FileCabinet cabinet in cabinets)
            //{
            //    try
            //    {
            //        List<Document> file_in_cabinet = GetAllFileInCabinet(cabinet.Id);
            //        if (file_in_cabinet.Count > 0)
            //            alldocument = alldocument.Concat(file_in_cabinet).ToList();
            //    }
            //    catch (Exception ex)
            //    {
            //        Logging.Log("Error get search in cabinet", ex);
            //    }

            //}

            return alldocument.ToList();

            return alldocument.ToList();
        }

        public string GetAllFileXML()
        {
            List<Document> documents = GetAllFile();
            Serializator<List<Document>> ser = new Serializator<List<Document>>();
            return ser.SerializationXMLString(documents);
        }

        public List<Document> Search(string field_name,string contractNo)
        {
            List<FileCabinet> cabinets = GetAllFileCabinets();
            if (cabinets == null)
            {
                Logging.Log("Error get cabinet list in action GetAllFile()");
                return null;
            }
            //ConcurrentBag<Document> alldocument = new ConcurrentBag<Document>();
            List<Document> alldocument = new List<Document>();
            object locker = new object();
            Parallel.ForEach(cabinets, cabinet =>
            {
                try
                {
                    //DocuWare docuWare = new DocuWare();
                    //docuWare.ServerUrl = ServerUrl;
                    //docuWare.ServerUrl = ServerUrl;
                    //docuWare.UseWindowsAuthentication =this.UseWindowsAuthentication;
                    //docuWare.UserName = UserName;
                    //docuWare.Password = Password;
                    //docuWare.ConectServer();
                    //List<Document> file_in_cabinet = docuWare.GetAllFileInCabinet(cabinet.Id).Where(d => ((d?.Fields?.Find(f => f.FieldName == field_name)?.Item as string) ?? "").IndexOf(contractNo, StringComparison.OrdinalIgnoreCase) >= 0)?.ToList();
                    //foreach (var doc in file_in_cabinet)
                    //{
                    //    alldocument.Add(doc);
                    //}
                    var dialogInfoItems = cabinet.GetDialogInfosFromSearchesRelation();
                    var dialog = dialogInfoItems.Dialog[0].GetDialogFromSelfRelation();

                    var q = new DialogExpression(){
                Operation = DialogExpressionOperation.And,
                Condition = new List<DialogExpressionCondition>()
                {
                    DialogExpressionCondition.Create(field_name, contractNo, contractNo )
                },
                Count = int.MaxValue,
                SortOrder = new List<SortedField> 
                { 
                    SortedField.Create("DWSTOREDATETIME", SortDirection.Desc)
                }};
                var queryResult = dialog.GetDocumentsResult(q);

                    lock (locker)
                    {
                        alldocument = alldocument.Concat(queryResult.Items.ToList()).ToList();
                    }
                    //docuWare.DisconectServer();
                }
                catch (Exception ex)
                {
                    Logging.Log("Error thread stop search", ex);
                }
            });
            //List<Document> alldocument = new List<Document>();
            //foreach (FileCabinet cabinet in cabinets)
            //{
            //    try
            //    {
            //        List<Document> file_in_cabinet = GetAllFileInCabinet(cabinet.Id).Where(d => ((d.Fields.Find(f => f.FieldName == field_name).Item as string)?? "").IndexOf(contractNo, StringComparison.OrdinalIgnoreCase)  >= 0)?.ToList();
            //        if (file_in_cabinet.Count > 0)
            //            alldocument = alldocument.Concat(file_in_cabinet).ToList();
            //    }
            //    catch(Exception ex)
            //    {
            //        Logging.Log("Error get search in cabinet",ex);
            //    }

            //}

            return alldocument.ToList();
        }

        public string SearchXML(string field_name, string contractNo)
        {
            List<Document> documents = Search(field_name,contractNo);
            Serializator<List<Document>> ser = new Serializator<List<Document>>();
            return ser.SerializationXMLString(documents);
        }

        public string DownloadDocumet(string cabinet_id, string file_id, string downloadPatch)
        {
            if (conector == null)
            {
                Logging.Log("Not conected to server cant download");
                return "";
            }
            try
            {
                Document document = conector.GetFromDocumentForDocumentAsync(int.Parse(file_id), cabinet_id).Result;
                var downloadResponse = document.PostToFileDownloadRelationForStreamAsync(
                    new FileDownload()
                    {
                        TargetFileType = FileDownloadType.Auto
                    }).Result;
                string save_url = downloadPatch + downloadResponse.GetFileName();
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
                return ex.ToString();
            }
        }

        public string GetDownloadDocumentUrl(string cabinet_id, string file_id)
        {

            if (conector == null)
            {
                Logging.Log("Not conected to server cant download");
                return "";
            }
            try
            {
                Document document = conector.GetFromDocumentForDocumentAsync(int.Parse(file_id), cabinet_id).Result;
                return document.FileDownloadRelationLink;
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

        public string GetLogFileLocation()
        {
            return Logging.getLogFileLocation();
        }
    }
}
