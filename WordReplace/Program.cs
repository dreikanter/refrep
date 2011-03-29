using System;
using System.IO;
using System.Linq;
using WordReplace.Extensions;
using WordReplace.References;

namespace WordReplace
{
    class Program
    {
        static void Main(string[] args)
        {
			var param = new Params(args, Console.Out);
			if (!param.Ready) return;

			var refFile = Path.GetFullPath(param.RefFile);
			if (!File.Exists(refFile))
			{
				LogWriteLine("Reference file does not exists: " + param.RefFile);
				return;
			}

			var srcFile = Path.GetFullPath(param.SourceFile);
			if (!File.Exists(srcFile))
			{
				LogWriteLine("Source document does not exists: " + param.SourceFile);
				return;
			}

			try
			{
				ProcessDoc(srcFile, param.DestFile.IsNullOrEmpty() ? 
					Utils.GetNewName(srcFile) : param.DestFile, refFile, param.Order);
			}
			catch (Exception ex)
			{
				LogWriteLine(ex.Message);
			}
        }

		private static void ProcessDoc(string srcFile, string destFile, string refFile, ReferenceOrder order)
		{
			LogWrite("Reading references from {0} ({1})... ".
				Fill(Path.GetFileName(refFile), refFile.GetFileSize().ToFormattedFileSize()));
			
			var refs = ReadReferences(refFile);
			
			if (refs.IsNullOrEmpty())
			{
				LogWriteLine("Got no references. Processing stopped.");
				return;
			}

			LogWriteLine("Done. Got {0} records.".Fill(refs.Count));
			LogWriteLine("Opening {0} ({1})...".Fill(Path.GetFileName(srcFile), srcFile.GetFileSize().ToFormattedFileSize()));

			using (var proc = new DocProcessor(srcFile, refs, order))
			{
				LogWriteLine("Found {0} reference groups; {1} bad IDs; {2} unknown IDs; {3} unknown tags".
					Fill(proc.Replacer.Replacements.Count, proc.Replacer.BadIds, proc.Replacer.UnknownIds, proc.Replacer.UnknownTags));

				if (proc.Replacer.BadIds.Any()) LogWriteLine("Bad IDs: " + proc.Replacer.BadIds.CommaSeparated());
				if (proc.Replacer.UnknownIds.Any()) LogWriteLine("Unknown IDs: " + proc.Replacer.UnknownIds.Cast<string>().CommaSeparated());
				if (proc.Replacer.UnknownTags.Any()) LogWriteLine("Unknown IDs: " + proc.Replacer.UnknownTags.CommaSeparated());

				proc.Process(destFile);
			}
		}

    	private static ReferenceCollection ReadReferences(string fileName)
		{
			try
			{
				using (var reader = new ReferenceReader(fileName))
				{
					return new ReferenceCollection(reader);
				}
			}
			catch (Exception ex)
			{
				LogWriteLine(ex.Message);
				return null;
			}
		}

		private static void LogWriteLine(string message)
		{
			Console.WriteLine(message);
		}

		private static void LogWrite(string message)
		{
			Console.Write(message);
		}
	}
}
