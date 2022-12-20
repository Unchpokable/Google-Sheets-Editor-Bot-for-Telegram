using Google.Apis.Sheets.v4.Data;
using System.Collections;
using System.Text;

namespace GSheetsEditor.Commands.Modules
{
    internal class SpreadsheetsCollection : IEnumerable<Spreadsheet>
    {
        public SpreadsheetsCollection()
        {
            _sheets = new List<Spreadsheet>();
            _selectedSpreadsheet = 0;
        }

        public SpreadsheetsCollection(IList<Spreadsheet> sheets, int selected) 
        {
            _sheets = sheets;
            _selectedSpreadsheet = selected;
        }

        private IList<Spreadsheet> _sheets;
        private int _selectedSpreadsheet;

        public Spreadsheet Current => _sheets[_selectedSpreadsheet];

        public void Add(Spreadsheet sheet)
        {
            _sheets.Add(sheet);
        }

        public void Remove(Spreadsheet sheet)
        {
            _sheets.Remove(sheet);
        }

        public Spreadsheet Next()
        {
            _selectedSpreadsheet++;
            if (_selectedSpreadsheet >= _sheets.Count)
            {
                _selectedSpreadsheet--;
            }

            return _sheets[_selectedSpreadsheet];
        }

        public Spreadsheet Previous()
        {
            _selectedSpreadsheet--;
            if (_selectedSpreadsheet < 0)
            {
                _selectedSpreadsheet++;
            }
            return _sheets[_selectedSpreadsheet];
        }

        public Spreadsheet SwitchTo(int selection)
        {
            if (selection < 0 || selection >= _sheets.Count)
                throw new ArgumentOutOfRangeException(nameof(selection));

            _selectedSpreadsheet = selection;
            return _sheets[_selectedSpreadsheet];
        }

        public override string ToString()
        {
            int index = 0;
            var sb = new StringBuilder();
            foreach (var item in _sheets)
            {
                sb.Append($"{index} : {item.Properties?.Title ?? "Untitled spreadsheet"}\n");
            }
            return sb.ToString();
        }

        public IEnumerator<Spreadsheet> GetEnumerator()
        {
            return _sheets.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_sheets).GetEnumerator();
        }
    }
}
