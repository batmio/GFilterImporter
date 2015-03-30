using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.ServiceModel.Syndication;
using System.Xml.Linq;
using CommandLine;
using CommandLine.Text;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace GFilterImporter
{
    class Program
    {
        // Command Line Options
        class Options
        {
            [Option('f', "file", Required = true, HelpText = "Input file to read.")]
            public string InputFile { get; set; }

            [Option('u', "user", Required = false, HelpText = "Specific User.")]
            public string UserName { get; set; }

            [Option('v', null, HelpText = "Print details during execution.")]
            public bool Verbose { get; set; }

            [HelpOption]
            public string GetUsage()
            {
                // this without using CommandLine.Text
                //  or using HelpText.AutoBuild
                var usage = new StringBuilder();
                usage.AppendLine("{Process.GetCurrentProcess()} 1.0");
                usage.AppendLine("-f for file");
                usage.AppendLine("-u for file");
                return usage.ToString();
            }
        }

        // MAIN
        static void Main(string[] args)
        {

            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                // Values are available here
                // Input File 
                ParseFilters(options.Verbose, options.InputFile, options.UserName);
            }

        } // end main


        ///
        public static String ParseFilters(bool verbose, string mailFilters, string user)
        {

            string mailFilter = mailFilters;
            if (File.Exists(@mailFilter))
            {
                // XML Document Loader
                XDocument doc = XDocument.Load(@mailFilter);

                if (verbose) Console.WriteLine("Parsing file {0}", mailFilter);
                
                // create our new list
                List<Entry> entries = new List<Entry>();

                foreach (XElement currentElement in doc.Root.Elements())
                {
                    if (currentElement.Name == "{http://www.w3.org/2005/Atom}entry")
                    {
                        IEnumerable<XElement> innerelements = currentElement.Descendants();
                        foreach (XElement innerCurrentElement in innerelements)
                        {
                            Entry entry = new Entry();

                            if (innerCurrentElement.Attribute("name") != null && innerCurrentElement.Attribute("name").Value == "from")
                            {
                                entry.From = innerCurrentElement.Attribute("value").Value.ToString();
                                //if (verbose) Console.WriteLine("From: {0}", innerCurrentElement.Attribute("value").Value);
                            }

                            if (innerCurrentElement.Attribute("name") != null && innerCurrentElement.Attribute("name").Value == "label")
                            {
                                // Here we take the label and change it to a folder namespace.
                                entry.Folder = innerCurrentElement.Attribute("value").Value.ToString();
                                //if (verbose) Console.WriteLine("Folder: {0}", innerCurrentElement.Attribute("value").Value);
                            }


                            // This will be useless soon.
                            entries.Add(entry);
                        }
                    }
                }
          
                //
                Console.ForegroundColor = ConsoleColor.Red;
                if (verbose) Console.WriteLine("Total Filters (Exchange Rules): " + entries.Count);
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                if (verbose) Console.WriteLine("The File {0} does not exist.", mailFilter);
                Console.ResetColor();

            }

            return null;
        }

        // Create Rule
        private void CreateExchangeRule(string folder, string email, string emailUser)
        {
            // Take the folder and the email name and parse it into a usable rule.



        }

        //Entry class
        class Entry
        {
            public string From;
            public string Folder;
        }

        // create the rules

    }
}