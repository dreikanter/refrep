﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using WordReplace.Extensions;

namespace WordReplace.References
{
	/// <summary>
	/// Класс, считывающий объекты Reference из таблицы Excel.
	/// </summary>
    public class ReferenceReader : IEnumerable<Reference>, IDisposable
    {
        private readonly string _workbookFileName;
        
        private readonly List<Reference> _references;
        
        private readonly Workbook _workBook;

        public ReferenceReader(string workbookFileName)
        {
            _workbookFileName = workbookFileName;
            _references = new List<Reference>();

            try
            {
                var excelApp = new Application();
                _workBook = excelApp.Workbooks.Open(workbookFileName);

                ExcelScanIntenal(_workBook);
            }
            catch(Exception ex)
            {
                throw new Exception("Error reading " + workbookFileName, ex);
            }
        }

        private void ExcelScanIntenal(Workbook workBook)
        {
            if (workBook == null) throw new ArgumentNullException("workBook");

            var numSheets = workBook.Sheets.Count;
            if (numSheets < 1) return;

            var sheet = (Worksheet) workBook.Sheets[1];
            var valueArray = (object[,]) sheet.UsedRange.Value[XlRangeValueDataType.xlRangeValueDefault];

            var rowsNum = valueArray.GetLength(0);
            var colsNum = valueArray.GetLength(1);

            if (rowsNum < 2 || colsNum < 1) return;

            var order = new List<ReferenceFields>(colsNum);
            var unknownFields = new List<string>();

            for (var i = 1; i <= colsNum; i++)
            {
                var fieldName = valueArray.GetValue(1, i).ToString();
                var field = fieldName.GetEnumValueOrDefault<ReferenceFields>();
                
                if (field == ReferenceFields.Unknown)
                {
                    unknownFields.Add(fieldName);
                }
                else
                {
                    order.Add(field);
                }
            }

            if (unknownFields.Any())
            {
                throw new Exception("Unknown fields found: " + unknownFields.CommaSeparated().InBrackets());
            }

            for (var row = 2; row < rowsNum; row++)
            {
                var values = new List<string>(colsNum);

                for (var col = 1; col <= colsNum; col++)
                {
                    var cell = valueArray.GetValue(row, col);
                    
                    if (cell == null)
                    {
                        values.Add(String.Empty);
                    }
                    else if (cell is DateTime)
                    {
                    	values.Add(((DateTime) cell).ToString(Constants.DateTimeFormat));
                    }
                    else
                    {
                        values.Add(cell.ToString().Trim());
                    }
                }

                _references.Add(new Reference(row, order, values));
            }
        }

        public IEnumerator<Reference> GetEnumerator()
        {
            return _references.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _references.GetEnumerator();
        }

        public void Dispose()
        {
            _workBook.Close(false, _workbookFileName, null);
            Marshal.ReleaseComObject(_workBook);
        }
    }
}
