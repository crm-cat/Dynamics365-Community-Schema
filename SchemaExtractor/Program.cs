using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dyn365CommunitySchema
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new CommandLineOptions();
            Parser.Default.ParseArguments(args, options);

            var schemaExtractor = new SchemaXmlExtractor(
                options.Entities.Split(','),
                options.ConnectionString,
                options.TargetFolder);

            schemaExtractor.Connect();
            schemaExtractor.Extract();

            Console.WriteLine("Extract Complete");

        }
    }
}
