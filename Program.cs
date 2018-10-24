using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Webritter.SharePointFileMover
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            log.Info("Programm started ");
            if (args.Count() != 1)
            {
                Console.WriteLine("Missing Parameter: optionXmlFile");
                log.Error("Missing Parameter: optionXmlFile");
                return;
            }

            if (args[0] == "sample")
            {
                RunOptions.GreateSampleXml("sample.xml");
                string message = "Sample xml file created: sample.xml";
                Console.WriteLine(message);
                log.Info(message);
                return;

            }
            if (args[0] == "sample2")
            {
                RunOptions.GreateSampleUndoXml("sample2.xml");
                string message = "Sample xml file created: sample2.xml";
                Console.WriteLine(message);
                log.Info(message);
                return;
            }

            string xmlFileName = args[0];
            if (!System.IO.File.Exists(xmlFileName))
            {
                Console.WriteLine("File does not exist: " + xmlFileName);
                log.Error("File does not exist: " + xmlFileName);
                return;
            }

            RunOptions options;
            try
            {
                options = RunOptions.LoadFromXMl(xmlFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Can't read and validate xmlFile: " + xmlFileName);
                log.Error("Can't read and validate xmlFile: " + xmlFileName);
                return;
            }


            if (string.IsNullOrEmpty(options.SiteUrl))
            {
                string message = "Missing SiteUrl in xmlFile: " + xmlFileName;
                Console.WriteLine(message);
                log.Error(message);
                return;
            }

            //Get instance of Authentication Manager  
            AuthenticationManager authenticationManager = new AuthenticationManager();
            //Create authentication array for site url,User Name,Password and Domain  
            try
            {
                SecureString password = GetSecureString(options.Password);
                //Create the client context  
                using (var ctx = authenticationManager.GetNetworkCredentialAuthenticatedContext(options.SiteUrl, options.Username, password, options.Domain))
                {
                    Site site = ctx.Site;
                    ctx.Load(site);
                    ctx.ExecuteQuery();
                    log.Info("Succesfully authenticated as " + options.Username);

                    List spList = ctx.Web.Lists.GetByTitle(options.LibraryName);
                    ctx.Load(spList);
                    ctx.ExecuteQuery();

                    log.Info("DocumentLibrary found with total " + spList.ItemCount + " docuements");

                    if (spList != null && spList.ItemCount > 0)
                    {
                        // build viewFields
                        string viewFields;
                        viewFields = "<ViewFields>";
                        if (!string.IsNullOrEmpty(options.StatusFieldName))
                        {
                            viewFields += "FieldRef Name='" + options.StatusFieldName + "' Type='Text' </FieldRef>";
                        }
                        viewFields += "</ViewFields>";

                        // build caml query
                        CamlQuery camlQuery = new CamlQuery();
                        camlQuery.ViewXml = "<View  Scope='Recursive'> " +
                                                "<Query>" +
                                                    options.CamlQuery +
                                                "</Query>" +
                                                viewFields +
                                            "</View>";

                        ListItemCollection listItems = spList.GetItems(camlQuery);
                        ctx.Load(listItems);
                        ctx.ExecuteQuery();

                        log.Info("found " + listItems.Count + " documents to check");


                        foreach (var item in listItems)
                        {
                            string filename = item["FileLeafRef"].ToString();
                           
                            log.Info("Checking '" + filename + "' ....");
                            bool skip = false;
                            List<object> fieldValues = new List<object>();
                            if (item.FileSystemObjectType == FileSystemObjectType.File)
                            {
                                File file = item.File;
                                ctx.Load(file);
                                ctx.ExecuteQuery();
                                string currentPath = file.ServerRelativeUrl.ToString();
                                string newPath = options.MoveTo;
                                if (!newPath.StartsWith("/"))
                                {
                                    // new path is a sub folder
                                    newPath = currentPath.Replace(filename, "") + options.MoveTo + "/" + filename;
                                }
                                else
                                {
                                    // new path is server relative
                                    newPath += "/" + filename;
                                }
                                file.MoveTo(newPath, MoveOperations.Overwrite);
                                ctx.ExecuteQuery();
                                log.Info("Moved ''" + item["FileLeafRef"] + "' to '" + options.MoveTo + "'");
                                if (!string.IsNullOrEmpty(options.StatusFieldName))
                                {
                                    // try to setup status 
                                    item[options.StatusFieldName] = options.StatusSuccessValue;
                                    item.Update();
                                    ctx.ExecuteQuery();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception : " + ex.Message);
            }


        }

        private static SecureString GetSecureString(string pwd)
        {
            var passWord = new SecureString();
            foreach (char c in pwd.ToCharArray()) passWord.AppendChar(c);
            return passWord;
        }

        private static bool IsDictionary(object o)
        {
            if (o == null) return false;
            return o is IDictionary &&
                   o.GetType().IsGenericType &&
                   o.GetType().GetGenericTypeDefinition().IsAssignableFrom(typeof(Dictionary<,>));
        }
    }
}
