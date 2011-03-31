using System;
using System.IO;
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

		public ReferenceReplacer Replacer { get { return _rep; } }

		private ListTemplate _listTemplate;

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
			FireMessage("Replacing references...");

			_doc.Activate();

			foreach (var pair in _rep.Replacements)
			{
				if (!pair.Value.Any()) continue;
				foreach (Range range in _doc.StoryRanges)
				{
					range.Find.Text = pair.Key;
					range.Find.Replacement.Text = pair.Value.GetOrderedRefNumList().InBrackets();
					range.Find.Wrap = WdFindWrap.wdFindContinue;
					range.Find.Execute(Replace: WdReplace.wdReplaceAll);
				}
			}
		}

		/// <summary>
		/// Генерация и вставка библиографического списка в документ, 
		/// по букмарку Constants.RefListBookmark.
		/// </summary>
		private void InsertRefList()
		{
			FireMessage("Generating bibliography list...");

			_doc.Activate();
			var range = GetRefListRange(Constants.RefListBookmark);
			range.Select();

			_word.Selection.Range.ListFormat.ApplyListTemplateWithLevel(GetListTemplate(), false, 
				WdListApplyTo.wdListApplyToWholeList, WdDefaultListBehavior.wdWord10ListBehavior);

			foreach(var reference in _rep.UsedReferences)
			{
				_word.Selection.TypeText(ReferenceCreator.GetReferenceText(reference));
				_word.Selection.TypeParagraph();
			}

			_word.Selection.TypeBackspace();
		}

		private void Save(string fileName)
		{
			_word.ActiveDocument.SaveAs(Path.GetFullPath(fileName));
			FireMessage("Document saved as {0}".Fill(fileName));
		}

		private Range GetRefListRange(string bookmark)
		{
			if(_doc == null) throw new InvalidOperationException();

			try
			{
				// ReSharper disable UseIndexedProperty
				return _doc.Bookmarks.get_Item(bookmark).Range;
				// ReSharper restore UseIndexedProperty
			}
			catch (Exception ex)
			{
				throw new Exception("Reference list bookmark ('{0}') not found".Fill(bookmark), ex);
			}
		}

		private ListTemplate GetListTemplate()
		{
			if (_listTemplate != null) return _listTemplate;

			_listTemplate = _word.ListGalleries[WdListGalleryType.wdNumberGallery].ListTemplates[1];

			var level = _listTemplate.ListLevels[1];

			level.NumberFormat = "%1.";
			level.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
			level.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
			level.NumberPosition = _word.InchesToPoints(0.25f);
			level.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
			level.TextPosition = _word.InchesToPoints(0.5f);
			level.TabPosition = (float) WdConstants.wdUndefined;
			level.ResetOnHigher = 0;
			level.StartAt = 1;

			level.Font.Bold = (int) WdConstants.wdUndefined;
			level.Font.Italic = (int) WdConstants.wdUndefined;
			level.Font.StrikeThrough = (int) WdConstants.wdUndefined;
			level.Font.Subscript = (int) WdConstants.wdUndefined;
			level.Font.Superscript = (int) WdConstants.wdUndefined;
			level.Font.Shadow = (int) WdConstants.wdUndefined;
			level.Font.Outline = (int) WdConstants.wdUndefined;
			level.Font.Emboss = (int) WdConstants.wdUndefined;
			level.Font.Engrave = (int) WdConstants.wdUndefined;
			level.Font.AllCaps = (int) WdConstants.wdUndefined;
			level.Font.Hidden = (int) WdConstants.wdUndefined;
			level.Font.Underline = WdUnderline.wdUnderlineNone;
			level.Font.Color = WdColor.wdColorAutomatic;
			level.Font.Size = (int) WdConstants.wdUndefined;
			level.Font.Animation = WdAnimation.wdAnimationNone;
			level.Font.DoubleStrikeThrough = (int) WdConstants.wdUndefined;

			level.LinkedStyle = String.Empty;

			_listTemplate.Name = "Bibliography reference list";

			return _listTemplate;
		}

		private void FireMessage(string message)
		{
			if (Message != null) Message(message);
		}

		public void Dispose()
		{
			_word.Quit(false);
		}
	}
}
