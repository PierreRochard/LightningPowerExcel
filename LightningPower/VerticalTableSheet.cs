using System.Collections.Generic;
using System.Linq;
using Google.Protobuf;
using Google.Protobuf.Reflection;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public class VerticalTableSheet<TMessageClass> where TMessageClass : IMessage
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;
        public TMessageClass Data;

        public int StartRow;
        private int _dataStartRow;
        public int StartColumn;
        public int EndColumn;
        public int EndRow;
        private readonly IList<FieldDescriptor> _fields;
        private readonly IReadOnlyCollection<string> _excludeList;

        public VerticalTableSheet(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor messageDescriptor, 
            IReadOnlyCollection<string> excludeList = default(List<string>))
        {
            Ws = ws;
            LApp = lApp;
            _fields = messageDescriptor.Fields.InDeclarationOrder();
            _excludeList = excludeList;
        }
        
        public void SetupVerticalTable(string tableName, int startRow = 2, int startColumn = 2)
        {
            StartRow = startRow;
            _dataStartRow = startRow + 1;
            StartColumn = startColumn;
            EndColumn = StartColumn + 1;

            if (_excludeList == null)
            {
                EndRow = startRow + _fields.Count;
            }
            else
            {
                EndRow = startRow + _fields.Count(f => !_excludeList.Any(f.Name.Contains));
            }

            var title = Ws.Cells[StartRow, StartColumn];
            title.Font.Italic = true;
            title.Value2 = tableName;

            var table = Ws.Range[Ws.Cells[_dataStartRow, StartColumn], Ws.Cells[EndRow, EndColumn]];
            Formatting.VerticalTable(table);

            var header = Ws.Range[Ws.Cells[_dataStartRow, StartColumn], Ws.Cells[EndRow, StartColumn]];
            Formatting.VerticalTableHeaderColumn(header);

            var data = Ws.Range[Ws.Cells[_dataStartRow, EndColumn], Ws.Cells[EndRow, EndColumn]];
            Formatting.VerticalTableDataColumn(data);

            var rowIndex = 0;
            foreach (var field in _fields)
            {
                if (_excludeList != null && _excludeList.Any(field.Name.Contains)) continue;

                var rowNumber = _dataStartRow + rowIndex;

                var headerCell = Ws.Cells[rowNumber, StartColumn];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;

                var rowRange = Ws.Range[Ws.Cells[rowNumber, StartColumn], Ws.Cells[rowNumber, EndColumn]];
                Formatting.VerticalTableRow(rowRange, rowNumber);

                rowIndex++;
            }

            title.Columns.AutoFit();
        }

        public void Clear()
        {
            var data = Ws.Range[Ws.Cells[_dataStartRow, EndColumn], Ws.Cells[EndRow, EndColumn]];
            data.ClearContents();
            Data = default(TMessageClass);
        }

        public void Update(TMessageClass newMessage)
        {
            var isCached = Data != null;
            if (isCached && Data.Equals(newMessage))
            {
                return;
            }
           
            if (!isCached)
            {
                Populate(newMessage);
            }
            else
            {
                Update(newMessage, Data);
            }
        }

        private static string GetValue(TMessageClass message, IFieldAccessor targetField)
        {
            return targetField.GetValue(message).ToString();
        }

        public void Populate(TMessageClass newMessage)
        {
            var rowIndex = 0;
            foreach (var field in _fields)
            {
                if (_excludeList != null && _excludeList.Any(field.Name.Contains)) continue;

                var rowNumber = _dataStartRow + rowIndex;
                var dataCell = Ws.Cells[rowNumber, EndColumn];
                var dataRow = Ws.Range[Ws.Cells[rowNumber, StartColumn], dataCell];
                Formatting.VerticalTableRow(dataRow, rowNumber);

                var newValue = GetValue(newMessage, field.Accessor);
                Utilities.AssignCellValue(newMessage, field, newValue, dataCell);
                Ws.Names.Add(field.Name, dataCell);
                rowIndex++;
            }
            Data = newMessage;
        }

        public void Update(TMessageClass newMessage, TMessageClass oldMessage)
        {
            for (var fieldIndex = 0; fieldIndex < _fields.Count; fieldIndex++)
            {
                var field = _fields[fieldIndex];
                var newValue = field.Accessor.GetValue(newMessage).ToString();
                var oldValue = field.Accessor.GetValue(oldMessage).ToString();
                if (oldValue == newValue) continue;

                var dataCell = Ws.Cells[_dataStartRow + fieldIndex, EndColumn];
                Utilities.AssignCellValue(newMessage, field, newValue, dataCell);
            }
            Data = newMessage;

        }
    }
}