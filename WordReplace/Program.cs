using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NDesk.Options;
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
				Console.WriteLine("Reference file does not exists: " + param.RefFile);
				return;
			}

			var srcFile = Path.GetFullPath(param.SourceFile);
			if (!File.Exists(srcFile))
			{
				Console.WriteLine("Source document does not exists: " + param.SourceFile);
				return;
			}

			Console.Write("Reading references from {0} ({1})... ".
				Fill(Path.GetFileName(refFile), refFile.GetFileSize().ToFormattedFileSize()));

			var e = File.Exists(refFile);
			var refs = ReadReferences(refFile);
			if (refs.IsNullOrEmpty()) return;
			Console.WriteLine("Done. Got {0} records.".Fill(refs.Count));

			Console.WriteLine("Opening {0} ({1})...".
				Fill(Path.GetFileName(srcFile), srcFile.GetFileSize().ToFormattedFileSize()));

			try
			{
				using (var proc = new DocProcessor(srcFile, refs, param.Order))
				{
					Console.WriteLine("Found {0} reference groups; {1} bad IDs; {2} unknown IDs; {3} unknown tags".
						Fill(proc.Replacer.Replacements.Count, proc.Replacer.BadIds, proc.Replacer.UnknownIds, proc.Replacer.UnknownTags));

					if (proc.Replacer.BadIds.Any()) Console.WriteLine("Bad IDs: " + proc.Replacer.BadIds.CommaSeparated());
					if (proc.Replacer.UnknownIds.Any()) Console.WriteLine("Unknown IDs: " + proc.Replacer.UnknownIds.CommaSeparated());
					if (proc.Replacer.UnknownTags.Any()) Console.WriteLine("Unknown IDs: " + proc.Replacer.UnknownTags.CommaSeparated());

					proc.Process(param.DestFile.IsNullOrEmpty() ? Utils.GetNewName(param.SourceFile) : param.DestFile);
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
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
				Console.WriteLine(ex.Message);
				return null;
			}
		}
    }
}
