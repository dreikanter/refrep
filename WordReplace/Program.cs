using System;
using System.Diagnostics;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WordReplace
{
    class Program
    {
        private static string GetExecPath()
        {
            var mainModule = Process.GetCurrentProcess().MainModule;
            return mainModule == null ? String.Empty : Path.GetDirectoryName(mainModule.FileName);
        }

        static void Main(string[] args)
        {
            var path = GetExecPath();
            object source = Path.Combine(path, "source.doc");
            object target = Path.Combine(path, "target.rtf");

            var word = new Word.Application { Visible = false };
            var missing = Type.Missing;


            var doc = word.Documents.Open(ref source, ref missing,
                     ref missing, ref missing, ref missing,
                     ref missing, ref missing, ref missing,
                     ref missing, ref missing, ref missing,
                     ref missing, ref missing, ref missing, ref missing);

            var text = doc.Content.Text;

            doc.Activate();

            // Loop through the StoryRanges (sections of the Word doc)
            foreach (Word.Range tmpRange in doc.StoryRanges)
            {
                // Set the text to find and replace
                tmpRange.Find.Text = "[#123]";
                tmpRange.Find.Replacement.Text = "]321[";

                // Set the Find.Wrap property to continue (so it doesn't
                // prompt the user or stop when it hits the end of
                // the section)
                tmpRange.Find.Wrap = Word.WdFindWrap.wdFindContinue;

                // Declare an object to pass as a parameter that sets
                // the Replace parameter to the "wdReplaceAll" enum
                object replaceAll = Word.WdReplace.wdReplaceAll;

                // Execute the Find and Replace -- notice that the
                // 11th parameter is the "replaceAll" enum object
                tmpRange.Find.Execute(ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref replaceAll,
                    ref missing, ref missing, ref missing, ref missing);
            }

//            object format = Word.WdSaveFormat.wdFormatUnicodeText;

            word.ActiveDocument.SaveAs(ref target, ref missing, // format
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing);

            word.Quit(ref missing, ref missing, ref missing);
        }
    }
}
