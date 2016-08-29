using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


    public class CommandLineOptions
    {

        [Option('f', "SourceFolder", HelpText = "Path to the output folder", Required = false)]
        public string TargetFolder { get; set; }

        [Option('c', "ConnectionString", HelpText = "Connection String", Required = true)]
        public string ConnectionString { get; set; }

        [Option('e', "Entities", HelpText = "Entities to extract", Required = false, DefaultValue="account,contact,customeraddress")]
        public string Entities { get; set; }


        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }

