using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordReplace.Extensions;
using WordReplace.References;

namespace WordReplace
{
	static class Program
    {
		[STAThread]
        static void Main(string[] args)
        {
			var param = new Params(args, Console.Out);
			if (!param.Ready) return;

			if (!File.Exists(param.RefFile))
			{
				LogWriteLine("Reference file does not exists: " + Path.GetFileName(param.RefFile));
				return;
			}

			if (!File.Exists(param.SourceFile))
			{
				LogWriteLine("Source document does not exists: " + Path.GetFileName(param.SourceFile));
				return;
			}

			try
			{
				ProcessDoc(param.SourceFile, param.DestFile.IsNullOrEmpty() ? 
					Utils.GetNewName(param.SourceFile) : param.DestFile, param.RefFile, param.Order);
			}
			catch (Exception ex)
			{
				LogWriteLine(ex.Message);
			}
        }

		///<summary>
		/// This method launches DocProcessor and handles document processing.
		/// </summary>
		/// <remarks>
		/// Each *FileName parameter is converted to full file path inside of this method. In other 
		/// case this may lead to file-not-found exception during Word or Excel interoperation (external 
		/// processes may have different current directories, that's why relative pathes may not work).
		/// </remarks>
		private static void ProcessDoc(string sourceFileName, string destFileName, string refFileName, ReferenceOrder order)
		{
			var refFile = Path.GetFullPath(refFileName);
			var srcFile = Path.GetFullPath(sourceFileName);
			var destFile = Path.GetFullPath(destFileName);

			LogWriteLine("{0} % {1} -> {2}".Fill(Path.GetFileName(sourceFileName), 
				Path.GetFileName(refFileName), Path.GetFileName(destFileName)));

			LogWrite("Reading references from {0} ({1})... ".
				Fill(Path.GetFileName(refFile), refFile.GetFileSize().ToFormattedFileSize()));
			
			var refs = ReadReferences(refFile);

			try
			{
				ValidateReferences(refs);
			}
			catch(Exception ex)
			{
				LogWriteLine("\n{0}\n{1}".Fill(ex.Message, "Processing stopped."));
				return;
			}

			LogWriteLine("Done. Got {0} records.".Fill(refs.Count));
			LogWriteLine("Opening {0} ({1})...".Fill(Path.GetFileName(srcFile), srcFile.GetFileSize().ToFormattedFileSize()));

			using (var proc = new DocProcessor(srcFile, refs, order))
			{
				proc.Message += OnMessage;
				ReportProcState(proc);

				try
				{
					proc.Process(destFile);
				}
				catch (Exception ex)
				{
					LogWriteLine("Error: " + ex.Message);
				}
			}
		}

		private static void ValidateReferences(IEnumerable<Reference> refs)
		{
			if (refs.IsNullOrEmpty())
			{
				throw new Exception("Got no references.");
			}

			var errorMessages = new List<string>();
			
			foreach (var reference in refs)
			{
				ICollection<string> errors;
				if (!ReferenceValidator.Validate(reference, out errors))
				{
					errorMessages.Add("Row {0}: {1}".Fill(reference.RowNum, errors.SemicolonSeparated()));
				}
			}

			if(errorMessages.Any())
			{
				throw new Exception("Invalid records data found:\n" + errorMessages.JoinWith("\n"));
			}
		}

		private static void ReportProcState(DocProcessor proc)
		{
			LogWriteLine("Found {0} reference groups; {1} bad IDs; {2} unknown IDs; {3} unknown tags".
				Fill(proc.Replacer.Replacements.Count, proc.Replacer.BadIds.Length,
					proc.Replacer.UnknownIds.Length, proc.Replacer.UnknownTags.Length));

			if (proc.Replacer.BadIds.Any())
			{
				LogWriteLine("Bad IDs: " + proc.Replacer.BadIds.CommaSeparated());
			}

			if (proc.Replacer.UnknownIds.Any())
			{
				LogWriteLine("Unknown IDs: " + proc.Replacer.UnknownIds.Cast<string>().CommaSeparated());
			}

			if (proc.Replacer.UnknownTags.Any())
			{
				LogWriteLine("Unknown Tags: " + proc.Replacer.UnknownTags.CommaSeparated());
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

		private static void OnMessage(string message)
		{
			LogWriteLine(message);
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
