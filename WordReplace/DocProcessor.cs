using System;
using System.Linq;
using Microsoft.Office.Interop.Word;
using WordReplace.Auxiliary;
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

		public ReferenceCollection References { get { return _refs; } }

		public ReferenceReplacer Replacer { get { return _rep; } }

		/// <summary>
		/// General purpose message event. Can be used for processing progress indication.
		/// </summary>
		public event MessageEventHandler Message;

		public DocProcessor(string fileName, ReferenceCollection refs, ReferenceOrder order)
		{
			if (fileName.IsNullOrBlank()) throw new ArgumentException("fileName");
			if (refs.IsNullOrEmpty()) throw new ArgumentException("refs");

			_word = new Application {Visible = false};
			_refs = refs;
			_order = order;

			try
			{
				_doc = _word.Documents.Open(fileName, ReadOnly: true);
			}
			catch(Exception ex)
			{
				throw new Exception("Error opening Word document", ex);
			}

			_rep = new ReferenceReplacer(_doc.Content.Text, ref _refs, _order);
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
				if (!pair.Value.Any()) continue;
				foreach (Range range in _doc.StoryRanges)
				{
					range.Find.Text = pair.Key;
					range.Find.Replacement.Text = "[{0}]".Fill(pair.Value.GetOrderedRefNumList());
					range.Find.Wrap = WdFindWrap.wdFindContinue;
					range.Find.Execute(Replace: WdReplace.wdReplaceAll);
				}
			}
		}

		private void InsertRefList()
		{
			FireMessage("Generating bibliography list...");	
		}

		private void Save(string fileName)
		{
			_word.ActiveDocument.SaveAs(fileName);
		}

		private void FireMessage(string message)
		{
			if (Message != null) Message(message);
		}

		public void Dispose()
		{
			_word.Quit(Type.Missing, Type.Missing, Type.Missing);
		}
	}
}
