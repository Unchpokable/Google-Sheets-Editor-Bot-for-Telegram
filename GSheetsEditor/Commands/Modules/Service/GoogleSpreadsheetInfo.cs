using Google.Apis.Sheets.v4.Data;
using System.Xml.Serialization;

namespace GSheetsEditor.Commands.Modules.Service
{
    [XmlType(TypeName = "GoogleSpreadsheetInfo")]
    public class GoogleSpreadsheetInfo
    {
        public GoogleSpreadsheetInfo() { }

        public GoogleSpreadsheetInfo(Spreadsheet origin)
        {
            Title = origin.Properties.Title;
            ID = origin.SpreadsheetId;
            Sheets = origin.Sheets.Select(sheet => sheet.Properties.Title).ToList();
        }
 
        [XmlAttribute("Title")] public string Title { get; set; }
        [XmlAttribute("Id")] public string ID { get; set; }

        [XmlAttribute("Sheets")] public List<string> Sheets { get; set; }
        public string WorkingSheet => Sheets[_workingSheet];

        private int _workingSheet;

        public void ChangeWorkingSheet(int selection)
        {
            if (selection < 0 || selection >= Sheets.Count)
                throw new ArgumentOutOfRangeException(nameof(selection));

            _workingSheet = selection;
        }

        public void NextSheet()
        {
            _workingSheet++;
            if (_workingSheet >= Sheets.Count)
                _workingSheet--;
        }

        public void PreviousSheet()
        {
            _workingSheet--;
            if(_workingSheet <= 0) 
                _workingSheet++;
        }
    }
}
