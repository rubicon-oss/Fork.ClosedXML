using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPrintAreas : IXLPrintAreas
    {
        private List<string> ranges = new List<string>();
        private XLWorksheet worksheet;

        public XLPrintAreas(XLWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public XLPrintAreas(XLPrintAreas defaultPrintAreas, XLWorksheet worksheet)
        {
            ranges = defaultPrintAreas.ranges.ToList();
            this.worksheet = worksheet;
        }

        public void Clear()
        {
            ranges.Clear();
        }

        public void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            ranges.Add(worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn).ToString());
        }

        public void AddExpression(string expression)
        {
            ranges.Add(expression);
        }

        public void Add(string rangeAddress)
        {
            ranges.Add(worksheet.Range(rangeAddress).ToString());
        }

        public void Add(string firstCellAddress, string lastCellAddress)
        {
            ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress).ToString());
        }

        public void Add(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress).ToString());
        }

        public IEnumerator<string> GetEnumerator()
        {
            return ranges.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
