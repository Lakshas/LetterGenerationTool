using RESTfulEngine.CSharpClient;
using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Configuration;

namespace WindwardLetterGeneration
{

    [Serializable(), ClassInterface(ClassInterfaceType.AutoDual), ComVisible(true)]
    public class WindWardGenerateLetters
    {
        public string errorMessage;
        public enum outputType : int { Docx, Pdf, Printer }

        //Parameters are all set in public as the library doesn't support App.config
        public string DEV = "http://10.1.54.16:8080/";
        public string QA = "http://10.1.54.14/";
        public string PROD = "http://10.1.54.16:8080/";
        public string DEVSERVER = "USNDCSSSQLD05";
        public string QASERVER = "USNDCSSSQLQ01";
        public string PRODSERVER = "USNDCSSSQLP01";
        public Uri MachineEnvironment()
        {
            Uri uRL = null;
            string eMacName = Environment.MachineName.ToString();

            if (eMacName == DEVSERVER)
            { uRL = new Uri(DEV); }
            else if (eMacName == QASERVER)
            { uRL = new Uri(QA); }
            else if (eMacName == PRODSERVER)
            { uRL = new Uri(PROD); }
            return uRL;
        }

        //[TestMethod] for finding out the versioning of the REST Engine
        public String Client_GetVersion()
        {
            string engine;
            RESTfulEngine.CSharpClient.Version v = Report.GetVersion(MachineEnvironment());
            engine = "Engine Version:"+ v.EngineVersion + "Service Version:"+ v.ServiceVersion;
            return engine;
        }

        public string GenerateLetters(string templateLoc, string outputLoc, outputType outputType, string server, string database, string autoTagSource, string[] paramFields, string[] paramValues, bool overwriteOuput = false)
        {
            string printStatus;
            

            if (File.Exists(outputLoc) && !overwriteOuput)
            {
                errorMessage = "Output file already exists(" + outputLoc + "). Default will not overwrite file;";
                return errorMessage;
            }

            switch (outputType)
            {
                case outputType.Docx: printStatus = PrintDocx(templateLoc, outputLoc, server, database, autoTagSource, paramFields, paramValues); break;
                case outputType.Pdf: printStatus = PrintPdf(templateLoc, outputLoc, server, database, autoTagSource, paramFields, paramValues); break;
                case outputType.Printer: printStatus = PrintPrinter(templateLoc, outputLoc, server, database, autoTagSource, paramFields, paramValues); break;
                default: errorMessage = "Improper Output Letter Type Selected"; printStatus = errorMessage; break;
            }

            return printStatus;
        }

        public string GenerateLettersQA(string templateLoc, string outputLoc, outputType outputType, string server, string database, string autoTagSource, string paramField, string paramValue, bool overwriteOuput = false)
        {
            string printStatus;

            if (File.Exists(outputLoc) && !overwriteOuput)
            {
                errorMessage = "Output file already exists(" + outputLoc + "). Default will not overwrite file;";
                return errorMessage;
            }

            switch (outputType)
            {
                case outputType.Docx: printStatus = PrintDocxWindward(templateLoc, outputLoc, server, database, autoTagSource, paramField, paramValue); break;
                //case outputType.Pdf: printStatus = PrintPdf(templateLoc, outputLoc, server, database, autoTagSource,paramFields, paramValues); break;
                //case outputType.Printer: printStatus = PrintPrinter(templateLoc, outputLoc, server, database, autoTagSource,paramFields, paramValues); break;
                default: errorMessage = "Improper Output Letter Type Selected"; printStatus = errorMessage; break;
            }

            return printStatus;
        }
        //[TestMethod] to test if the configuration file is working as expected
        public string GetInfo()
        {
            Uri uRL = null;
            string eMacName = Environment.MachineName.ToString();
            if (eMacName == ConfigurationManager.AppSettings["DEVServer"].ToString())
            { uRL = new Uri(ConfigurationManager.AppSettings["DEV"]); }
            else if (eMacName == ConfigurationManager.AppSettings["QAServer"].ToString())
            { uRL = new Uri(ConfigurationManager.AppSettings["QA"]); }
            else if (eMacName == ConfigurationManager.AppSettings["PRODServer"].ToString())
            { uRL = new Uri(ConfigurationManager.AppSettings["PROD"]); }
            return uRL.ToString();
        }

        public string Test_Object(object collection)
        {
            List<TemplateVariable> _collection = new List<TemplateVariable>();
            TemplateVariable singleton = new TemplateVariable();
            _collection = List< TemplateVariable >();



        }

        private string PrintDocx(string templateLoc, string outputLoc, string source, string catalog,string autoTagSource, string[] paramFields, string[] paramValues)
        {
            try
            {
                var ds = new AdoDataSource("System.Data.SqlClient", "Data Source = " + source + "; Integrated Security = True; Initial Catalog = " + catalog);

                ds.Variables = new List<TemplateVariable>();

                for (int i = 0; i < paramFields.Length; i++)
                {
                    ds.Variables.Add(new TemplateVariable() { Name = paramFields[i], Value = paramValues[i] });
                }

                var dataSources = new Dictionary<string, DataSource>()
                {
                {autoTagSource, ds}
                };

                using (var templateFile = File.OpenRead(templateLoc))
                {
                    using (var outputFile = File.Create(outputLoc))
                    {
                        var report = new ReportDocx(MachineEnvironment(), templateFile, outputFile);
                        report.Process(dataSources);
                    }
                }

                return "0";
            }

            catch (Exception e)
            {
                errorMessage = e.ToString();
                return errorMessage;
            }

        }

        private string PrintPdf(string templateLoc, string outputLoc, string source, string catalog, string autoTagSource,string[] paramFields, string[] paramValues)
        {
            try
            {

                var ds = new AdoDataSource("System.Data.SqlClient", "Data Source = " + source + "; Integrated Security = True; Initial Catalog = " + catalog);

                ds.Variables = new List<TemplateVariable>();
                
                for(int i = 0; i < paramFields.Length; i++)
                {
                    ds.Variables.Add(new TemplateVariable() { Name = paramFields[i], Value = paramValues[i] });
                }

                var dataSources = new Dictionary<string, DataSource>()
                {
                {autoTagSource, ds}
                };

                using (var templateFile = File.OpenRead(templateLoc))
                {
                    using (var outputFile = File.Create(outputLoc))
                    {
                        var report = new ReportPdf(MachineEnvironment(), templateFile, outputFile);
                        report.Process(dataSources);
                    }
                }

                return "0";
            }

            catch (Exception e)
            {
                errorMessage = e.ToString();
                return errorMessage;
            }
        }

        private string PrintPrinter(string templateLoc, string outputLoc, string source, string catalog, string autoTagSource, string[] paramFields, string[] paramValues)
        {
            try
            {
                var ds = new AdoDataSource("System.Data.SqlClient", "Data Source = " + source + "; Integrated Security = True; Initial Catalog = " + catalog);

                ds.Variables = new List<TemplateVariable>();

                for (int i = 0; i < paramFields.Length; i++)
                {
                    ds.Variables.Add(new TemplateVariable() { Name = paramFields[i], Value = paramValues[i] });
                }

                var dataSources = new Dictionary<string, DataSource>()
                {
                {autoTagSource, ds}
                };

                using (var templateFile = File.OpenRead(templateLoc))
                {
                    var report = new ReportPrinter(MachineEnvironment(), templateFile, outputLoc);
                    report.Process(dataSources);
                }

                return "0";
            }

            catch (Exception e)
            {
                errorMessage = e.ToString();
                return errorMessage;
            }
        }


        private string PrintDocxWindward(string templateLoc, string outputLoc, string source, string catalog, string autoTagSource,string paramField, string paramValue)
        {
            try
            {

                var ds = new AdoDataSource("System.Data.SqlClient", "Data Source = " + source + "; Integrated Security = True; Initial Catalog = " + catalog);

                ds.Variables = new List<TemplateVariable>();
                
                ds.Variables.Add(new TemplateVariable() { Name = paramField, Value = paramValue });
                
                var dataSources = new Dictionary<string, DataSource>()
                {
                {autoTagSource, ds}
                };

                using (var templateFile = File.OpenRead(templateLoc))
                {
                    using (var outputFile = File.Create(outputLoc))
                    {
                        var report = new ReportDocx(MachineEnvironment(), templateFile, outputFile);
                        report.Process(dataSources);
                    }
                }

                return "0";
            }

            catch (Exception e)
            {
                errorMessage = e.ToString();
                return errorMessage;
            }

        }
    }
}
