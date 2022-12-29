using Google.Apis.Sheets.v4.Data;
using System.Collections;
using System.Text;
using System.Xml.Serialization;

namespace GSheetsEditor.Commands.Modules.Service
{
    [XmlType(TypeName = "SpreadsheetsCollection")]
    public class SpreadsheetsCollection : IEnumerable<GoogleSpreadsheetInfo>
    {
        public SpreadsheetsCollection()
        {
            Sheets = new List<GoogleSpreadsheetInfo>();
            _selectedSpreadsheet = 0;
        }

        public SpreadsheetsCollection(IList<GoogleSpreadsheetInfo> sheets, int selected)
        {
            Sheets = sheets;
            _selectedSpreadsheet = selected;
        }

        [XmlArray("Sheets")] public IList<GoogleSpreadsheetInfo> Sheets { get; init; }
        private int _selectedSpreadsheet;

        public GoogleSpreadsheetInfo Current => Sheets[_selectedSpreadsheet];

        public void Add(GoogleSpreadsheetInfo sheet)
        {
            Sheets.Add(sheet);
        }

        public void Remove(GoogleSpreadsheetInfo sheet)
        {
            Sheets.Remove(sheet);
        }

        public GoogleSpreadsheetInfo Next()
        {
            _selectedSpreadsheet++;
            if (_selectedSpreadsheet >= Sheets.Count)
            {
                _selectedSpreadsheet--;
            }

            return Sheets[_selectedSpreadsheet];
        }

        public GoogleSpreadsheetInfo Previous()
        {
            _selectedSpreadsheet--;
            if (_selectedSpreadsheet < 0)
            {
                _selectedSpreadsheet++;
            }
            return Sheets[_selectedSpreadsheet];
        }

        public GoogleSpreadsheetInfo SwitchTo(int selection)
        {
            if (selection < 0 || selection >= Sheets.Count)
                throw new ArgumentOutOfRangeException(nameof(selection));

            _selectedSpreadsheet = selection;
            return Sheets[_selectedSpreadsheet];
        }

        public override string ToString()
        {
            int index = 0;
            var sb = new StringBuilder();
            foreach (var item in Sheets)
            {
                sb.Append($"{index} : {item.Title ?? "Untitled spreadsheet"}\n");
            }
            return sb.ToString();
        }

        public IEnumerator<GoogleSpreadsheetInfo> GetEnumerator()
        {
            return Sheets.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)Sheets).GetEnumerator();
        }
    }
}
