using System;
using System.Linq;
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
				_doc = _word.Documents.Open(fileName, ref _missing, ref _missing, ref _missing, 
					ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, 
					ref _missing, ref _missing, ref _missing, ref _missing, ref _missing);
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
			_doc.Activate();

			foreach (var pair in _rep.Replacements)
			{
				foreach (Range range in _doc.StoryRanges)
				{
					range.Find.Text = pair.Key;
					range.Find.Replacement.Text = "[{0}]".Fill(pair.Value.Cast<string>().CommaSeparatedNb());
					range.Find.Wrap = WdFindWrap.wdFindContinue;

					object replaceAll = WdReplace.wdReplaceAll;

					range.Find.Execute(ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, 
						ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref replaceAll, 
						ref _missing, ref _missing, ref _missing, ref _missing);
				}
			}
		}

		private void InsertRefList()
		{
			
		}

		private void Save(string fileName)
		{
			_word.ActiveDocument.SaveAs(fileName, ref _missing, ref _missing, ref _missing, 
				ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, 
				ref _missing, ref _missing, ref _missing, ref _missing, ref _missing, ref _missing);
		}

		public void Dispose()
		{
			_word.Quit(ref _missing, ref _missing, ref _missing);
		}
	}
}
