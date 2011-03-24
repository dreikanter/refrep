using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using WordReplace.Extensions;

namespace WordReplace.References
{
    public class ReferenceReader : IEnumerable<Reference>, IDisposable
    {
        private readonly string _workbookFileName;
        
        private readonly List<Reference> _references;
        
        private readonly Workbook _workBook;

        public delegate void ErrorHandler(string message);

        public event ErrorHandler Error;

        public ReferenceReader(string workbookFileName, ErrorHandler errorHandler)
        {
            _workbookFileName = workbookFileName;
            _references = new List<Reference>();
            Error += errorHandler;

            try
            {
                var excelApp = new Application();
                _workBook = excelApp.Workbooks.Open(workbookFileName,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing);

                ExcelScanIntenal(_workBook);
            }
            catch(Exception ex)
            {
                InvokeError(ex.Message);
            }
        }

        private void ExcelScanIntenal(Workbook workBookIn)
        {
            if (workBookIn == null) throw new ArgumentNullException("workBookIn");

            var numSheets = workBookIn.Sheets.Count;
            if (numSheets < 1) return;

            var sheet = (Worksheet) workBookIn.Sheets[1];
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
                throw new Exception("Unknown fields found: [{0}]".Fill(unknownFields.CommaSeparated()));
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
                        values.Add(((DateTime)cell).ToString(Constants.DateTimeFormat));
                    }
                    else
                    {
                        values.Add(cell.ToString());
                    }
                }

                _references.Add(new Reference(order, values));
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

        public void InvokeError(string message)
        {
            if (Error != null) Error(message);
        }
    }
}
