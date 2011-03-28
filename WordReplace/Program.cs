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
using Word = Microsoft.Office.Interop.Word;

namespace WordReplace
{
    class Program
    {
        private static IEnumerable<Reference> GetReferences(string fileName)
        {
            var refPath = Path.Combine(GetExecPath(), fileName);
            using (var reader = new ReferenceReader(refPath, Console.WriteLine))
            {
                foreach (var r in reader) yield return r;
            }
        }

        private static IEnumerable<Reference> GetUsedReferences(IEnumerable<Reference> allReferences, IEnumerable<int> usedIds, ReferenceOrder order)
        {
            var refDict = new Dictionary<int, Reference>(allReferences.Count());
            foreach (var reference in allReferences)
            {
                if (refDict.ContainsKey(reference.Id))
                {
                    throw new Exception("Duplicate reference ID used: '{0}'".Fill(reference.Id));
                }

                refDict.Add(reference.Id, reference);
            }

            var usedReferences = new List<Reference>();
            var unknownRefIds = new HashSet<int>();

            foreach (var usedId in usedIds)
            {
                if(!refDict.ContainsKey(usedId))
                {
                    unknownRefIds.Add(usedId);
                    continue;
                }

                var reference = refDict[usedId];
                if(!usedReferences.Contains(reference)) usedReferences.Add(reference);
            }

            if (unknownRefIds.Any())
            {
                throw new Exception("Unknown reference ID used: {0}".
                    Fill(unknownRefIds.Cast<String>().CommaSeparated()));
            }

            if (order == ReferenceOrder.Alpha)
            {
                usedReferences.Sort((x, y) => x.CompareTo(y));
            }

            var refCnt = 1;
            foreach (var usedReference in usedReferences)
            {
                usedReference.RefNum = refCnt++;
            }

            return usedReferences;
        }

        private static string GetExecPath()
        {
            var mainModule = Process.GetCurrentProcess().MainModule;
            return mainModule == null ? String.Empty : Path.GetDirectoryName(mainModule.FileName);
        }

        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage: greet [OPTIONS]+ message");
            Console.WriteLine("Greet a list of individuals with an optional message.");
            Console.WriteLine("If no message is specified, a generic greeting is used.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        private static void SearchIds(string text, out IEnumerable<int> usedIds, out IDictionary<string, int[]> refMatches)
        {
            var re = new Regex(@"\[\[([\d\s,]+)\]\]");
            var ids = new List<int>();
            var matches = new Dictionary<string, int[]>();

            foreach (Match match in re.Matches(text))
            {
                if (match.Groups.Count < 2 || match.Groups[1].Captures.Count < 1) continue;
                var matchIds = new List<int>();

                foreach (var value in match.Groups[1].Captures[0].Value.Split(new[] {Constants.RefIdDelimiter}))
                {
                    var refId = Int32.Parse(value, NumberStyles.Integer);
                    if (!ids.Contains(refId)) ids.Add(refId);
                    matchIds.Add(refId);
                }

                var matchString = match.Groups[0].Captures[0].Value;
                if (!matches.ContainsKey(matchString)) matches.Add(matchString, matchIds.ToArray());
            }

            usedIds = ids;
            refMatches = matches;
        }

		public static string ToReferenceText(IEnumerable<Reference> references)
		{
			if (references == null || !references.Any()) return String.Empty;
			return String.Format("[{0}]", String.Join(", ", references.Select(r => r.RefNum.ToString()).ToArray()));
		}

		public static void TestRefReplacing()
		{
			var refs = new ReferenceCollection
			           	{
							new Reference{Id = 1, RefNum=11, Tag="tag1"},
							new Reference{Id = 2, RefNum=12, Tag="tag2"},
							new Reference{Id = 3, RefNum=13, Tag="tag3"},
							new Reference{Id = 4, RefNum=14, Tag="таг1"},
							new Reference{Id = 5, RefNum=15, Tag="таг2"},
							new Reference{Id = 6, RefNum=16, Tag="таг3"},
							new Reference{Id = 7, RefNum=17, Tag="tag4"},
						};

			var text = File.ReadAllText("source.txt");
			var replacer = new ReferenceReplacer(text, ref refs, ReferenceOrder.Mention);
			replacer.UsedReferences.Enumerate();

			var sb = new StringBuilder(text);
			foreach (var repl in replacer.Replacements)
			{
				sb.Replace(repl.Key, ToReferenceText(repl.Value));
			}
		}

        static void Main(string[] args)
        {
            //var param = new Params(args, Console.Out);




            //return;

            //var path = GetExecPath();

            //object source = Path.Combine(path, "source.doc");
            //object target = Path.Combine(path, "target.doc");

            //var word = new Word.Application { Visible = false };
            //var missing = Type.Missing;

            //var doc = word.Documents.Open(ref source, ref missing,
            //         ref missing, ref missing, ref missing,
            //         ref missing, ref missing, ref missing,
            //         ref missing, ref missing, ref missing,
            //         ref missing, ref missing, ref missing, ref missing);

            //IEnumerable<int> ids;
            //IDictionary<string, int[]> matches;
            //SearchIds(doc.Content.Text, out ids, out matches);

//            doc.Activate();

//            // Loop through the StoryRanges (sections of the Word doc)
//            foreach (Word.Range tmpRange in doc.StoryRanges)
//            {
//                tmpRange.Find.Text = "[#123]";
//                tmpRange.Find.Replacement.Text = "]321[";
//                tmpRange.Find.Wrap = Word.WdFindWrap.wdFindContinue;
//                object replaceAll = Word.WdReplace.wdReplaceAll;

//                tmpRange.Find.Execute(ref missing, ref missing, ref missing,
//                    ref missing, ref missing, ref missing, ref missing,
//                    ref missing, ref missing, ref missing, ref replaceAll,
//                    ref missing, ref missing, ref missing, ref missing);
//            }

////            object format = Word.WdSaveFormat.wdFormatUnicodeText;

//            word.ActiveDocument.SaveAs(ref target, ref missing, // format
//                        ref missing, ref missing, ref missing,
//                        ref missing, ref missing, ref missing,
//                        ref missing, ref missing, ref missing,
//                        ref missing, ref missing, ref missing,
//                        ref missing, ref missing);

//            word.Quit(ref missing, ref missing, ref missing);
        }
    }
}
