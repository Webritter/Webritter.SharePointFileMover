using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Webritter.SharePointFileMover
{
    public class RunOptions
    {
        public int Id { get; set; }
        public bool Enabled { get; set; }
        public string Domain { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string SiteUrl { get; set; }
        public string LibraryName { get; set; }
        public string CamlQuery { get; set; }

        public string MoveTo { get; set; }

        public string StatusFieldName { get; set; }
        public string StatusSuccessValue { get; set; }


        public static RunOptions LoadFromXMl(string xmlFileName)
        {
            // Now we can read the serialized book ...  
            System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(RunOptions));
            System.IO.StreamReader file = new System.IO.StreamReader(xmlFileName);
            RunOptions result = (RunOptions)reader.Deserialize(file);
            file.Close();
            return result;
        }

        // file load and save
        public void SaveAsXml(string xmlFileName)
        {

            var writer = new System.Xml.Serialization.XmlSerializer(typeof(RunOptions));
            var wfile = new System.IO.StreamWriter(xmlFileName);
            writer.Serialize(wfile, this);
            wfile.Close();
        }

        public static void GreateSampleXml(string filename)
        {
            RunOptions sample = new RunOptions()
            {
                SiteUrl = "http://sharepoint.webritter.tk/sites/dev",
                Domain = "",
                Username = "webritter",
                Password = "secret",
                LibraryName = "Documents",
                CamlQuery = "<Where><Eq><FieldRef Name='Status' /><Value Type='Text'>Ready to Archive</Value></Eq></Where>",
                MoveTo = "Archive",
                StatusFieldName = "Status",
                StatusSuccessValue = "Archived"
            };
            sample.SaveAsXml(filename);
        }

        public static void GreateSampleUndoXml(string filename)
        {
            RunOptions sample = new RunOptions()
            {
                SiteUrl = "http://sharepoint.webritter.tk/sites/dev",
                Domain = "",
                Username = "webritter",
                Password = "secret",
                LibraryName = "Documents",
                CamlQuery = "<Where><And><Eq><FieldRef Name='Status' /><Value Type='Text'>Archived</Value></Eq><Eq><FieldRef Name='FileDirRef' /><Value Type='Text'>/sites/dev/Shared Documents/Archive</Value></Eq></And></Where>",
                MoveTo = "/sites/dev/Shared Documents",
                StatusFieldName = "Status",
                StatusSuccessValue = "Draft"
            };
            sample.SaveAsXml(filename);
        }
    }
}
