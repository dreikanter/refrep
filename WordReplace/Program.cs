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
