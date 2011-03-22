using System;
using System.Collections.Generic;
using System.IO;
using NDesk.Options;
using WordReplace.Extensions;

namespace WordReplace
{
    public class Params
    {
        private readonly TextWriter _outputWriter;

        public string SourceFile { get; private set; }
        
        public string DestFile { get; private set; }
        
        public string RefFile { get; private set; }
        
        public ReferenceOrder Order { get; private set; }

        public bool Ready { get; private set; }

        public Params(IEnumerable<string> args, TextWriter outputWriter)
        {
            _outputWriter = outputWriter;
            var showHelp = false;
            var p = new OptionSet
                        {
                            { "s|source=", "Source Word document (*.doc[x] file)", v => SourceFile = v },
                            { "d|destination=", "Processed document name", v => DestFile = v },
                            { "r|references=", "References spreadsheet (*.xls[x] file)", v => RefFile= v },
                            { "o|order=", "Sort order for references (alpha|mention)", v => Order = v.GetEnumValueOrDefault<ReferenceOrder>() },
                            { "h|help",  "Show this message and exit", v => showHelp = (v != null) }
                        };

            try
            {
                p.Parse(args);
            }
            catch (OptionException e)
            {
                WriteMessage("WordReplace: ");
                WriteMessage(e.Message);
                WriteMessage("Try `greet --help' for more information.");
                Ready = false;
                return;
            }

            if (showHelp)
            {
                ShowHelp(p);
                Ready = false;
                return;
            }

            var noSource = SourceFile.IsNullOrBlank();
            var noRef = RefFile.IsNullOrBlank();

            if (noSource || noRef)
            {
                if (noSource) WriteMessage("Source file not specified");
                if (noRef) WriteMessage("Reference file not specified");
                Ready = false;
                return;
            }

            if (DestFile.IsNullOrBlank()) DestFile = GetDestinationFileName(SourceFile);

            Ready = true;
        }

        private void ShowHelp(OptionSet p)
        {
            WriteMessage("Usage: ref [OPTIONS]+");
            WriteMessage("");
            WriteMessage("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        private static string GetDestinationFileName(string sourceFile)
        {
            return Path.Combine(Path.GetDirectoryName(sourceFile), "{0}-ready{1}".
                Fill(Path.GetFileNameWithoutExtension(sourceFile), Path.GetExtension(sourceFile)));
        }
       
        public void WriteMessage(string message)
        {
            if (_outputWriter != null) _outputWriter.WriteLine(message);
        }
    }
}