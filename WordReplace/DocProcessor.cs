using System;
using Microsoft.Office.Interop.Word;
using WordReplace.Extensions;
using WordReplace.References;

namespace WordReplace
{
	/// <summary>
	/// Word Document processor.
	/// </summary>
	public class DocProcessor : IDisposable
	{
		private readonly Application _word;
		
		private readonly ReferenceCollection _refs;
		
		private readonly ReferenceOrder _order;

		private readonly ReferenceReplacer _rep;
		
		private readonly Document _doc;

		private object _missing = Type.Missing;

		public ReferenceCollection References { get { return _refs; } }

		public ReferenceReplacer Replacer { get { return _rep; } }

		public DocProcessor(string fileName, ReferenceCollection refs, ReferenceOrder order)
		{
			if (fileName.IsNullOrBlank()) throw new ArgumentException("fileName");
			if (refs.IsNullOrEmpty()) throw new ArgumentException("refs");

			_word = new Application {Visible = false};
			_refs = refs;
			_order = order;

			try
			{
				_doc = _word.Documents.Open(fileName, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing);
			}
			catch(Exception ex)
			{
				throw new Exception("Error opening Word document", ex);
			}

			_rep = new ReferenceReplacer(_doc.Content.Text, ref _refs, order);
		}

		public void Process(string saveTo)
		{
			ReplaceReferences();
			InsertRefList();
			Save(saveTo);
		}

		private void ReplaceReferences()
		{
			//doc.Activate();

			//// Loop through the StoryRanges (sections of the Word doc)
			//foreach (Word.Range tmpRange in doc.StoryRanges)
			//{
			//    tmpRange.Find.Text = "[#123]";
			//    tmpRange.Find.Replacement.Text = "]321[";
			//    tmpRange.Find.Wrap = Word.WdFindWrap.wdFindContinue;
			//    object replaceAll = Word.WdReplace.wdReplaceAll;

			//    tmpRange.Find.Execute(ref missing, ref missing, ref missing,
			//        ref missing, ref missing, ref missing, ref missing,
			//        ref missing, ref missing, ref missing, ref replaceAll,
			//        ref missing, ref missing, ref missing, ref missing);
			//}

			////            object format = Word.WdSaveFormat.wdFormatUnicodeText;
		}

		private void InsertRefList()
		{
			
		}

		private void Save(string fileName)
		{
			_word.ActiveDocument.SaveAs(fileName, ref _missing, // format
						ref _missing, ref _missing, ref _missing,
						ref _missing, ref _missing, ref _missing,
						ref _missing, ref _missing, ref _missing,
						ref _missing, ref _missing, ref _missing,
						ref _missing, ref _missing);
		}

		public void Dispose()
		{
			_word.Quit(ref _missing, ref _missing, ref _missing);
		}
	}
}
