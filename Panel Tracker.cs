//
// -*- coding: utf-8 -*-
//
//Created on Wed Apr 18 09:35:19 2018
//
//@author: JBradley
//

using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;

using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.SharePoint.Client;
using System.Security;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Text;
using System.IO;

namespace Panel_Tracker
{
    /// <summary>
    /// Dimension Selected Objects
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class Command : Autodesk.Revit.UI.IExternalCommand
    {
        static Dictionary<string, Panel> SharepointPull(string xlFileLink, XLGuide guide, out List<string> panFilter, out Dictionary<string, Panel> panList)
        {
            panFilter = new List<string>();
            panList = new Dictionary<string, Panel>();

            ClientContext ctx = new ClientContext("https://islandcompanies.sharepoint.com/sites/SiteDeliveries/");
            SecureString pw = new SecureString();
            foreach (char a in "Michael1994") pw.AppendChar(a);
            ctx.Credentials = new SharePointOnlineCredentials("Jbradley@islandcompanies.com", pw);
            //ctx.Credentials = new NetworkCredential("jbradley@islandcompanies.com", "Michael1994", "islandcompanies");

            Web web = ctx.Web;
            ctx.Load(web, w => w.Title);

            getXLFileName(xlFileLink, out string xlFileName);
            getSharepointTree(xlFileLink, out List<String> sharepointTree);

            try
            {
                Folder sourceFolder = web.GetFolderByServerRelativeUrl(sharepointTree.Aggregate((i, j) => i + "/" + j));
                ctx.Load(ctx.Web);
                ctx.Load(sourceFolder);
                ctx.ExecuteQuery();

                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='Recursive'>
                                        <Query>
                                            <Where>
                                                <Eq>
                                                    <FieldRef Name='FileLeafRef'></FieldRef>
                                                    <Value Type='Text'>" + xlFileName + @"</Value>
                                                </Eq>
                                            </Where>
                                        </Query>
                                    </View>";
                //string srvRel = "sites/SiteDeliveries/Shared%20Documents/Packlists/01-000-000%20Panel%20Tracker/";
                ListItemCollection listItems = ctx.Web.Lists.GetByTitle("Documents").GetItems(camlQuery);
                ctx.Load(listItems);
                ctx.ExecuteQuery();

                Microsoft.SharePoint.Client.File file = listItems[0].File;
                ctx.Load(file);
                ctx.ExecuteQuery();

                FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, file.ServerRelativeUrl);
                ctx.ExecuteQuery();

                string filePath = @"C:\Documents\Panel Tracker\" + file.Name;
                using (var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
                {
                    fileInfo.Stream.CopyTo(fileStream);
                }

                SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true);
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = document.WorkbookPart.Workbook;
                string sheetName = "Panel Tracker";
                WorkbookPart workbookpt = document.WorkbookPart;
                Sheet sheet = workbook.Descendants<Sheet>().Where(S => S.Name == sheetName).FirstOrDefault();

                WorksheetPart wsPart = (WorksheetPart)(workbookpt.GetPartById(sheet.Id));

                bool run = true;
                int iter = 0;
                while (run == true)
                {
                    Panel p = new Panel();

                    bool hasVal = false;
                    bool record = false;
                    string panName = "";
                    foreach (KeyValuePair<string, string> entry in guide.Info)
                    {
                        try
                        {
                            Cell cell = new Cell();
                            cell = GetCell(document, sheetName, entry.Value.ToString() + guide.Row.ToString(), out cell);
                            hasVal = true;
                            string v = cell.CellValue.InnerText;
                            if (entry.Key == "Panel Name")
                            {
                                p.Parameters.Add(entry.Key, v);
                                panName = v;
                            }
                            else
                            {
                                record = true;
                                string searchval = v.Split(')').Last();
                                p.Parameters.Add(entry.Key, searchval);
                            }
                        }
                        catch //(Exception ex)
                        {
                            p.Parameters.Add(entry.Key, null);
                            //Console.WriteLine(ex.Message);
                        }

                        iter++;
                    }
                    if (!hasVal)
                    {
                        run = false;
                    }
                    else
                    {
                        if (record)
                        {
                            panFilter.Add(panName);
                        }
                        panList.Add(panName, p);
                    }
                    guide.Row++;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            //Console.ReadKey();

            return panList;
        }

        public class Panel
        {
            public string Name { get; set; }
            public Dictionary<string, string> Parameters = new Dictionary<string, string>();

        }

        public class XLGuide
        {
            public int Row = 5;
            public Dictionary<string, string> Info = new Dictionary<string, string>();

        }

        public static Cell GetCell(SpreadsheetDocument document, string sheetName, string addressName, out Cell cell)
        {
            WorkbookPart wbPart = document.WorkbookPart;
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

            cell = (Cell)wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();


            string cellVal;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                //it's a shared string so use the cell inner text as the index into the 
                //shared strings table
                var stringId = Convert.ToInt32(cell.InnerText);
                cellVal = wbPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(stringId).InnerText;
            }
            else
            {
                //it's NOT a shared string, use the value directly
                cellVal = cell.InnerText;
            }

            cell.CellValue = new CellValue(cellVal);
            return cell;
        }

        public static string getXLFileName(string xlFileLink, out string xlFileName)
        {
            int xlNameStart = xlFileLink.LastIndexOf("/");
            int xlNameEnd = xlFileLink.LastIndexOf(".xlsx");

            xlFileName = xlFileLink.Substring(xlNameStart + 1, xlNameEnd - xlNameStart + 4);
            xlFileName = xlFileName.Replace("%20", " ");

            return xlFileName;
        }

        public static List<string> getSharepointTree(string xlFileLink, out List<string> sharepointTree)
        {
            sharepointTree = new List<string>();
            string[] splitSharepointLink = xlFileLink.Split('/');

            bool start = false;
            foreach (string s in splitSharepointLink)
            {
                if (s.Contains("Documents"))
                {
                    start = true;
                }
                if (start)
                {
                    sharepointTree.Add(s);
                }
            }

            sharepointTree.RemoveAt(sharepointTree.Count - 1);

            return sharepointTree;
        }

        public static List<Folder> getAllItems(ClientContext ctx, List DocumentsList, out List<Folder> allFolders)
        {
            CamlQuery camlQuery = new CamlQuery();
            camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View Scope='Recursive'> " +
                        "<Query>" +
                        "</Query>" +
                        "</View>";
            ListItemCollection listItems = DocumentsList.GetItems(camlQuery);
            ctx.Load(listItems);
            ctx.ExecuteQuery();
            foreach (ListItem item in listItems)
            {
                ctx.Load(item.File);
                ctx.ExecuteQuery();
                //Console.WriteLine(item.File.Name);
            }

            allFolders = new List<Folder>();
            return allFolders;
        }

        public static List<Element> GetAssemblies(Document doc, UIDocument uidoc, out List<Element> col)
        {
            //assemblies = new List<AssemblyInstance>();
            col = new FilteredElementCollector(doc).OfClass(typeof(AssemblyInstance)).ToList();

            return col;
        }

        public static void PanelUpdate(Document doc, UIDocument uidoc, List<string> panFilter, Dictionary<string, Panel> panList, List<Element> assemblies, bool useParameter, string parameterName)
        {
            if (useParameter == false)
            {
                foreach (AssemblyInstance a in assemblies)
                {
                    if (panFilter.Contains(a.Name))
                    {
                        foreach (KeyValuePair<string, string> entry in panList[a.Name].Parameters)
                        {
                            try
                            {
                                Autodesk.Revit.DB.Parameter p = a.LookupParameter(entry.Key);

                                p.Set(entry.Value);
                            }
                            catch
                            {

                            }
                        }
                    }
                }
            }
            else
            {
                foreach (AssemblyInstance a in assemblies)
                {
                    string pName = a.LookupParameter(parameterName).AsString();
                    if (panFilter.Contains(pName))
                    {
                        foreach (KeyValuePair<string, string> entry in panList[pName].Parameters)
                        {
                            try
                            {
                                Autodesk.Revit.DB.Parameter p = a.LookupParameter(entry.Key);

                                p.Set(entry.Value);
                            }
                            catch
                            {

                            }
                        }
                    }
                }
            }
        }

        public static void DocsUpload_(string bearer, string referenceURL, string filePath)
        {
            string hub_id = "";
            string account_id = "";
            string project_id = "";
            string object_id = "";
            string objectsub_id = "";
            bool fileExists = false;
            string lineage = "";

            int pFrom = referenceURL.IndexOf("projects/") + "projects/".Length;
            int pTo = referenceURL.LastIndexOf("/folders");

            string project_ref_id = referenceURL.Substring(pFrom, pTo - pFrom);

            pFrom = referenceURL.IndexOf("folders/") + "folders/".Length;
            pTo = referenceURL.LastIndexOf("/detail");

            string folder_id = referenceURL.Substring(pFrom, pTo - pFrom);

            string fileName = Path.GetFileName(filePath);

            string base_domain = "https://developer.api.autodesk.com/";

            // GET HUB 

            try
            {
                var webRequest = System.Net.WebRequest.Create(base_domain + "project/v1/hubs");
                if (webRequest != null)
                {
                    webRequest.Method = "GET";
                    webRequest.Timeout = 20000;
                    webRequest.ContentType = "application/json";
                    webRequest.Headers.Add("Authorization", bearer);
                    using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                    {
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                        {
                            var jsonResponse = sr.ReadToEnd();
                            dynamic responseBody = JObject.Parse(jsonResponse);
                            foreach (var index in responseBody.data)
                            {
                                if ((index.id as JValue).Value.ToString().Contains("b."))
                                {
                                    hub_id = (index.id as JValue).Value.ToString();
                                    account_id = (index.id as JValue).Value.ToString().Split('.')[1];
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

            // GET PROJECT

            try
            {
                var webRequest = System.Net.WebRequest.Create(base_domain + "project/v1/hubs/" + hub_id + "/projects");
                if (webRequest != null)
                {
                    webRequest.Method = "GET";
                    webRequest.Timeout = 20000;
                    webRequest.ContentType = "application/json";
                    webRequest.Headers.Add("Authorization", bearer);
                    using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                    {
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                        {
                            var jsonResponse = sr.ReadToEnd();
                            dynamic responseBody1 = JObject.Parse(jsonResponse);
                            foreach (var index1 in responseBody1.data)
                            {
                                if ((index1.id as JValue).Value.ToString().Contains(project_ref_id))
                                {
                                    project_id = (index1.id as JValue).Value.ToString();
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

            // GET CONTENTS

            try
            {
                var webRequest = System.Net.WebRequest.Create(base_domain + "data/v1/projects/" + project_id + "/folders/" + folder_id + "/contents");
                if (webRequest != null)
                {
                    webRequest.Method = "GET";
                    webRequest.Timeout = 20000;
                    webRequest.ContentType = "application/json";
                    webRequest.Headers.Add("Authorization", bearer);
                    using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                    {
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                        {
                            var jsonResponse = sr.ReadToEnd();
                            dynamic responseBody1 = JObject.Parse(jsonResponse);
                            foreach (var index in responseBody1.data)
                            {
                                if ((index.attributes.extension.data.sourceFileName as JValue).Value.ToString() == fileName
                                    && (index.attributes.displayName as JValue).Value.ToString() == fileName)
                                {
                                    fileExists = true;
                                    lineage = index.id;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

            // CREATE FILE BUCKET

            try
            {
                var webRequest = System.Net.WebRequest.Create(base_domain + "data/v1/projects/" + project_id + "/storage");
                if (webRequest != null)
                {
                    webRequest.Method = "POST";
                    webRequest.Timeout = 20000;
                    webRequest.ContentType = "application/vnd.api+json";
                    webRequest.Headers.Add("Authorization", bearer);

                    string postData = "{\n\"jsonapi\": { \"version\": \"1.0\" },\n\"data\": {\n\"type\": \"objects\",\n\"attributes\": {\n\"name\": \"" + fileName + "\"\n},\n\"relationships\": {\n\"target\": {\n\"data\": { \"type\": \"folders\", \"id\": \"" + folder_id + "\" }\n}\n}\n}\n}";
                    ASCIIEncoding encoding = new ASCIIEncoding();
                    byte[] byte1 = encoding.GetBytes(postData);
                    webRequest.ContentLength = byte1.Length;

                    using (var streamWriter = webRequest.GetRequestStream())
                    {
                        streamWriter.Write(byte1, 0, byte1.Length);
                    }

                    using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                    {
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                        {
                            var jsonResponse = sr.ReadToEnd();
                            dynamic responseBody3 = JObject.Parse(jsonResponse);
                            object_id = (responseBody3.data.id as JValue).Value.ToString();
                            objectsub_id = object_id.Substring(object_id.LastIndexOf('/'), object_id.Length - object_id.LastIndexOf('/'));
                            //MessageBox.Show(object_id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

            // UPLOAD FILE

            try
            {
                var webRequest = System.Net.WebRequest.Create(base_domain + "oss/v2/buckets/wip.dm.prod/objects/" + objectsub_id);
                if (webRequest != null)
                {
                    webRequest.Method = "PUT";
                    webRequest.Timeout = 20000;
                    webRequest.ContentType = "application/vnd.api+json";
                    webRequest.Headers.Add("Authorization", bearer);

                    byte[] byte1 = System.IO.File.ReadAllBytes(filePath);
                    webRequest.ContentLength = byte1.Length;

                    using (var streamWriter = webRequest.GetRequestStream())
                    {
                        streamWriter.Write(byte1, 0, byte1.Length);
                    }

                    using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                    {
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                        {
                            var jsonResponse = sr.ReadToEnd();
                            dynamic responseBody4 = JObject.Parse(jsonResponse);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

            if (!fileExists)
            {
                // VERSION FILE

                try
                {
                    var webRequest = System.Net.WebRequest.Create(base_domain + "data/v1/projects/" + project_id + "/items");
                    if (webRequest != null)
                    {
                        webRequest.Method = "POST";
                        webRequest.Timeout = 20000;
                        webRequest.ContentType = "application/vnd.api+json";
                        webRequest.Headers.Add("Authorization", bearer);

                        string postData = "{\n    \"jsonapi\": { \"version\": \"1.0\" },\n    \"data\": {\n      \"type\": \"items\",\n      \"attributes\": {\n        \"displayName\": \"" + fileName + "\",\n        \"extension\": {\n          \"type\": \"items:autodesk.bim360:File\",\n          \"version\": \"1.0\"\n        }\n      },\n      \"relationships\": {\n        \"tip\": {\n          \"data\": {\n            \"type\": \"versions\", \"id\": \"1\"\n          }\n        },\n        \"parent\": {\n          \"data\": {\n            \"type\": \"folders\",\n            \"id\": \"" + folder_id + "\"\n          }\n        }\n      }\n    },\n    \"included\": [\n      {\n        \"type\": \"versions\",\n        \"id\": \"1\",\n        \"attributes\": {\n          \"name\": \"" + fileName + "\",\n          \"extension\": {\n            \"type\": \"versions:autodesk.bim360:File\",\n            \"version\": \"1.0\"\n          }\n        },\n        \"relationships\": {\n          \"storage\": {\n            \"data\": {\n              \"type\": \"objects\",\n              \"id\": \"" + object_id + "\"\n            }\n          }\n        }\n      }\n    ]\n  }";
                        ASCIIEncoding encoding = new ASCIIEncoding();
                        byte[] byte1 = encoding.GetBytes(postData);
                        webRequest.ContentLength = byte1.Length;

                        using (var streamWriter = webRequest.GetRequestStream())
                        {
                            streamWriter.Write(byte1, 0, byte1.Length);
                        }

                        using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                        {
                            using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                            {
                                var jsonResponse = sr.ReadToEnd();
                                dynamic responseBody5 = JObject.Parse(jsonResponse);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {

                // VERSION UP

                try
                {
                    var webRequest = System.Net.WebRequest.Create(base_domain + "data/v1/projects/" + project_id + "/versions");
                    if (webRequest != null)
                    {
                        webRequest.Method = "POST";
                        webRequest.Timeout = 20000;
                        webRequest.ContentType = "application/vnd.api+json";
                        webRequest.Headers.Add("Authorization", bearer);

                        string postData = "{\n   \"jsonapi\": { \"version\": \"1.0\" },\n   \"data\": {\n      \"type\": \"versions\",\n      \"attributes\": {\n         \"name\": \"" + fileName + "\",\n         \"extension\": { \"type\": \"versions:autodesk.bim360:File\", \"version\": \"1.0\"}\n      },\n      \"relationships\": {\n         \"item\": { \"data\": { \"type\": \"items\", \"id\": \"" + lineage + "\" } },\n         \"storage\": { \"data\": { \"type\": \"objects\", \"id\": \"" + object_id + "\" } }\n      }\n   }\n}";
                        ASCIIEncoding encoding = new ASCIIEncoding();
                        byte[] byte1 = encoding.GetBytes(postData);
                        webRequest.ContentLength = byte1.Length;

                        using (var streamWriter = webRequest.GetRequestStream())
                        {
                            streamWriter.Write(byte1, 0, byte1.Length);
                        }

                        using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                        {
                            using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                            {
                                var jsonResponse = sr.ReadToEnd();
                                dynamic responseBody5 = JObject.Parse(jsonResponse);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public List<string> categoryparameters(Document doc, BuiltInCategory cat)
        {
            if (cat == null)
            {
                return null;
            }
            List<string> parameters = new List<string>();
            using (Transaction tr = new Transaction(doc, "make_schedule"))
            {
                tr.Start();
                // Create schedule
                ViewSchedule vs = ViewSchedule.CreateSchedule(doc, Category.GetCategory(doc, cat).Id);
                doc.Regenerate();

                // Find schedulable fields
                foreach (SchedulableField sField in vs.Definition.GetSchedulableFields())
                {
                    if (sField.FieldType != ScheduleFieldType.ElementType) continue;
                    parameters.Add(sField.GetName(doc));
                }
                tr.RollBack();
            }
            return parameters;
        }

        public Autodesk.Revit.UI.Result Execute(
        Autodesk.Revit.UI.ExternalCommandData commandData,
        ref string message, Autodesk.Revit.DB.ElementSet elementSet)
        {
            Autodesk.Revit.UI.UIApplication revit = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Transaction trans = new Transaction(revit.ActiveUIDocument.Document, "Panel Tracker");
            ElementSet collection = new ElementSet();

            Dictionary<string, List<String>> categories = new Dictionary<string, List<string>>();
            foreach (BuiltInCategory c in System.Enum.GetValues(typeof(BuiltInCategory)))
            {
                string cName = Category.GetCategory(doc, c).Name;
                categories.Add(cName, categoryparameters(doc, c));
            }

            string filePath = @"C:\Documents\Panel Tracker\tracker.txt";

            string xlFileLink = "";
            string xlnamerow = "";
            string xlnamecol = "";
            string xlstatusrow = "";
            string xlstatuscol = "";
            string bimDocsLink = "";
            bool useParameter = false;
            string parameterName = "";

            if (System.IO.File.Exists(filePath) && !(System.Windows.Forms.Control.ModifierKeys == Keys.N))
            {
                // Open the file to read from.
                using (System.IO.StreamReader sr = System.IO.File.OpenText(filePath))
                {
                    string s = "";
                    List<String> ls = new List<String>();
                    while ((s = sr.ReadLine()) != null)
                    {
                        ls.Add(s);
                    }

                    xlFileLink = ls[0];
                    xlnamerow = ls[1];
                    xlnamecol = ls[2];
                    xlstatusrow = ls[3];
                    xlstatuscol = ls[4];
                    bimDocsLink = ls[5];

                    if (ls[6] == "1") { useParameter = true; } else { useParameter = false; }
                    parameterName = ls[7];
                }
            }
            else
            {
                mLocalForms = new LocalForms();
                mLocalForms.CreateForm();

                //Example with direct call to sub menu.                
                bool displayForm = true;
                while (displayForm)
                {
                    if (mLocalForms.trackerForm == null)
                    {
                        mLocalForms.trackerForm = new Tracker();
                    }
                    if (!formOpen)
                    {
                        mLocalForms.trackerForm.ShowDialog();
                        formOpen = true;
                    }
                    if (mLocalForms.trackerForm.exit)
                    {
                        displayForm = false;
                    }
                }
                if (mLocalForms.trackerForm.filled)
                {
                    bimDocsLink = mLocalForms.trackerForm.bimLink;
                    xlFileLink = mLocalForms.trackerForm.sharepointLink;
                    xlnamerow = mLocalForms.trackerForm.namerow;
                    xlnamecol = mLocalForms.trackerForm.namecolumn;
                    xlstatusrow = mLocalForms.trackerForm.statusrow;
                    xlstatuscol = mLocalForms.trackerForm.statuscolumn;

                    useParameter = mLocalForms.trackerForm.useParameter;
                    parameterName = mLocalForms.trackerForm.parameter;

                    // Create a file to write to.
                    using (System.IO.StreamWriter sw = System.IO.File.CreateText(filePath))
                    {
                        sw.WriteLine(xlFileLink);
                        sw.WriteLine(xlnamerow);
                        sw.WriteLine(xlnamecol);
                        sw.WriteLine(xlstatusrow);
                        sw.WriteLine(xlstatuscol);
                        sw.WriteLine(bimDocsLink);

                        if (useParameter == true) { sw.WriteLine("1"); } else { sw.WriteLine("0"); }
                        sw.WriteLine(parameterName);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("UNFILLED");
                }
            }


            XLGuide guide = new XLGuide();
            guide.Info.Add("Panel Name", xlnamecol);
            guide.Info.Add("Status", xlstatuscol);

            try
            {
                guide.Row = Int32.Parse(xlnamerow);
            }
            catch (FormatException)
            {
                guide.Row = 0;
            }

            GetAssemblies(doc, uidoc, out List<Element> assemblies);

            assemParameters = new List<string>();

            foreach (Autodesk.Revit.DB.Parameter p in assemblies[0].Parameters)
            {
                assemParameters.Add(p.Definition.Name);
            }

            bool error = false;
            try
            {
                SharepointPull(xlFileLink, guide, out List<string> panFilter, out Dictionary<string, Panel> panList);

                //####################################################
                //#"Start" the transaction
                trans.Start();

                PanelUpdate(doc, uidoc, panFilter, panList, assemblies, useParameter, parameterName);

                //# "End" the transaction
                trans.Commit();
                //####################################################

                mLocalForms2 = new LocalForms2();
                mLocalForms2.CreateForm();
                mLocalForms2.oauthForm.ShowDialog();

                string fileName = "";
                string oldFileName = doc.PathName;

                string docTitle = doc.Title;

                if (docTitle.Substring(doc.Title.Length - 4) == ".rvt")
                {
                    docTitle = docTitle.Substring(0, doc.Title.Length - 4);
                }
                if (docTitle.Contains("Panel_Tracker"))
                {
                    fileName = @"C:\Documents\Panel Tracker\" + docTitle + ".rvt";
                }
                else
                {
                    fileName = @"C:\Documents\Panel Tracker\" + docTitle + "Panel_Tracker.rvt";
                }
                SaveAsOptions options = new SaveAsOptions();
                options.OverwriteExistingFile = true;
                doc.SaveAs(fileName, options);

                doc.SaveAs(oldFileName, options);

                DocsUpload_(mLocalForms2.oauthForm.bearer, bimDocsLink, fileName);
            }
            catch (Exception e)
            {
                // if revit threw an exception, try to catch it
                TaskDialog.Show("Error", e.Message);
                error = true;
                return Autodesk.Revit.UI.Result.Failed;
            }
            finally
            {
                // if revit threw an exception, display error and return failed
                if (error)
                {
                    TaskDialog.Show("Error", "Panel Tracker failed.");
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        private LocalForms mLocalForms;
        public static bool formOpen = false;

        private LocalForms2 mLocalForms2;
        public static bool formOpen2 = false;

        public static List<string> assemParameters { get; set; }
        public static Dictionary<string, List<string>> categories { get; set; }

    }

    public class LocalForms
    {

        public Tracker trackerForm
        {
            set; get;
        }

        public LocalForms()
        {

        }

        public void CreateForm()
        {
            trackerForm = new Tracker();
        }
    }

    public class LocalForms2
    {

        public OAuth oauthForm
        {
            set; get;
        }

        public LocalForms2()
        {

        }

        public void CreateForm()
        {
            oauthForm = new OAuth();
        }
    }
}
