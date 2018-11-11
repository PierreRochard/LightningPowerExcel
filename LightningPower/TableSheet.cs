using System;
using System.Collections.Generic;
using System.Linq;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public class TableSheet<TMessageClass> where TMessageClass : IMessage
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;
        
        public int StartRow;
        public int DataStartRow;
        public int HeaderRow;
        public int StartColumn;
        public int EndColumn;
        public int EndRow;

        public IList<FieldDescriptor> Fields;
        public Dictionary<object, TMessageClass> DisplayData;
        public List<TMessageClass> DataList;
        public List<TMessageClass> DisplayDataList;
        public Range Title;

        private readonly List<string> _wideColumns;
        private readonly IFieldAccessor _uniqueKeyField;
        private readonly string _uniqueKeyName;
        private readonly int _limit;
        private readonly string _sortColumn;
        private readonly bool _sortAscending;
        private readonly IFieldAccessor _uniqueNestedField;
        private readonly List<Tuple<IFieldAccessor, IFieldAccessor>> _nestedFields;

        public TableSheet(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor messageDescriptor, string uniqueKeyName, 
            bool nestedData = false, int limit = 0, string sortColumn = null, bool sortAscending = true)
        {
            _nestedFields = new List<Tuple<IFieldAccessor, IFieldAccessor>>();
            _uniqueKeyName = uniqueKeyName;
            Ws = ws;
            LApp = lApp;
            DisplayData = new Dictionary<object, TMessageClass>();
            Fields = messageDescriptor.Fields.InDeclarationOrder()
                .Where(f => f.FieldType != FieldType.Message || !nestedData).ToList();

            var messageFields = messageDescriptor.Fields.InDeclarationOrder()
                .Where(f => f.FieldType == FieldType.Message && nestedData && !f.IsRepeated && !f.IsMap).ToList();
            foreach (var parentField in messageFields)
            {
                var childFields = parentField.MessageType.Fields.InDeclarationOrder();
                Fields = Fields.Concat(childFields).ToArray();
                foreach (var childField in childFields)
                {
                    var nf = new Tuple<IFieldAccessor, IFieldAccessor>(parentField.Accessor, childField.Accessor);
                    if (!parentField.IsMap &&!parentField.IsRepeated) _nestedFields.Add(nf);
                    if (childField.Name != _uniqueKeyName) continue;
                    _uniqueKeyField = childField.Accessor;
                    _uniqueNestedField = parentField.Accessor;
                }
            }

            if (_uniqueKeyField == null)
            {
                foreach (var field in Fields)
                {
                    if (field.Name == _uniqueKeyName) _uniqueKeyField = field.Accessor;
                }
            }

            _limit = limit;
            _sortColumn = sortColumn;
            _sortAscending = sortAscending;
            _wideColumns = new List<string>
            {
                "pub_key",
                "remote_pubkey",
                "remote_pub_key",
                "remote_node_pub",
                "channel_point",
                "pending_htlcs",
                "closing_tx_hash",
                "closing_txid",
                "chain_hash",
                "payment_preimage",
                "payment_hash",
                "path",
                "tx_hash",
                "block_hash",
                "dest_addresses"
            };
        }

        public void SetupTable(string tableName, int rowCount = 3, int startRow = 2, int startColumn = 2)
        {
            StartRow = startRow;
            HeaderRow = startRow + 1;
            DataStartRow = HeaderRow + 1;
            StartColumn = startColumn;
            EndColumn = StartColumn + Fields.Count - 1;
            EndRow = HeaderRow + rowCount;


            Title = Ws.Cells[StartRow, StartColumn];
            Title.Value2 = tableName;

            var header = Ws.Range[Ws.Cells[HeaderRow, StartColumn], Ws.Cells[HeaderRow, EndColumn]];
            Formatting.TableHeaderRow(header);

            var data = Ws.Range[Ws.Cells[HeaderRow + 1, StartColumn], Ws.Cells[EndRow, EndColumn]];
            Formatting.TableDataColumn(data, false);

            Ws.Columns.AutoFit();
            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var columnNumber = StartColumn + fieldIndex;
                var headerCell = Ws.Cells[HeaderRow, columnNumber];
                var field = Fields[fieldIndex];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;

                if (field.IsRepeated && field.FieldType != FieldType.Message)
                {
                    Ws.Columns[columnNumber].ColumnWidth = 100;
                }

                var isWide = _wideColumns != null && _wideColumns.Any(field.Name.Contains);
                if (!isWide) continue;
                Formatting.WideTableColumn(Ws.Range[Ws.Cells[1, StartColumn], Ws.Cells[100, EndColumn]]);
            }

            for (var rowI = 0; rowI < rowCount; rowI++)
            {
                var rowNumber = rowI + DataStartRow;
                var rowRange = Ws.Range[Ws.Cells[rowNumber, StartColumn], Ws.Cells[rowNumber, EndColumn]];
                Formatting.TableDataRow(rowRange, rowNumber % 2 == 0);
            }

            Formatting.TableTitle(Title);
        }

        private object GetUniqueKey(TMessageClass message)
        {
            if (_uniqueNestedField == null)
            {
                return _uniqueKeyField.GetValue(message);
            }

            return _uniqueKeyField.GetValue((IMessage) _uniqueNestedField.GetValue(message));
        }

        private string GetValue(TMessageClass message, IFieldAccessor targetField)
        {
            foreach (var nestedField in _nestedFields)
            {
                var parentField = nestedField.Item1;
                var childField = nestedField.Item2;
                if (childField.Descriptor.Name == targetField.Descriptor.Name)
                {
                    return Utilities.GetValue(childField.Descriptor, (IMessage) parentField.GetValue(message));
                }
            }
            return Utilities.GetValue(targetField.Descriptor, message);
        }
        
        public void Update(RepeatedField<TMessageClass> data)
        {
            DataList = data.ToList();
            DisplayDataList = DataList;
            if (_sortColumn != null)
            {
                var prop = typeof(TMessageClass).GetProperty(_sortColumn);
                if (prop != null)
                {
                    DisplayDataList = _sortAscending 
                        ? DisplayDataList.OrderBy(d => prop.GetValue(d, null)).ToList() 
                        : DisplayDataList.OrderByDescending(d => prop.GetValue(d, null)).ToList();
                }
            }

            if (_limit > 0)
            {
                DisplayDataList = DisplayDataList.Take(10).ToList();
            }
            
            foreach (var newMessage in DisplayDataList)
            {
                var uniqueKey = GetUniqueKey(newMessage);
                var isCached = DisplayData.TryGetValue(uniqueKey, out var cachedMessage);
                if (isCached && cachedMessage.Equals(newMessage))
                {
                    continue;
                }

                DisplayData[uniqueKey] = newMessage;

                if (!isCached)
                {
                    AppendRow(newMessage);
                }
                else
                {
                    UpdateRow(newMessage, cachedMessage);
                }
            }

            foreach (var cachedUniqueKey in DisplayData.Keys)
            {
                var result = DisplayDataList.FirstOrDefault(newMessage => GetUniqueKey(newMessage).ToString() == cachedUniqueKey.ToString());
                if (result == null)
                {
                    RemoveRow(cachedUniqueKey);
                }
            }

            try
            {
                Ws.Columns.AutoFit();

            }
            catch (Exception)
            {
                return;
            }
            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var field = Fields[fieldIndex];
                var columnNumber = StartColumn + fieldIndex;
                var isWide = _wideColumns != null && _wideColumns.Any(field.Name.Contains);
                if (!isWide) continue;
                Formatting.WideTableColumn(Ws.Range[Ws.Cells[1, columnNumber], Ws.Cells[1, columnNumber]]);
            }

            Formatting.TableTitle(Title);
            EndRow = GetLastRow();
            Utilities.RemoveLoadingMark(Ws);
        }

        private void RemoveRow(object uniqueKey)
        {
            var rowNumber = GetRow(uniqueKey);
            RemoveRow(rowNumber);
        }

        public void RemoveRow(int rowNumber)
        {
            if (rowNumber == 0) return;
            var range = Ws.Range[Ws.Cells[rowNumber, StartColumn], Ws.Cells[rowNumber, EndColumn]];
            range.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        private void AppendRow(TMessageClass newMessage)
        {
            var lastRow = GetLastRow();
            Formatting.TableDataRow(Ws.Range[Ws.Cells[lastRow, StartColumn], Ws.Cells[lastRow, EndColumn]], lastRow % 2 == 0);
            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var field = Fields[fieldIndex];
                var columnNumber = StartColumn + fieldIndex;
                var dataCell = Ws.Cells[lastRow, columnNumber];
                var newValue = GetValue(newMessage, field.Accessor);
                Utilities.AssignCellValue(newMessage, field, newValue, dataCell);
                var isWide = _wideColumns != null && _wideColumns.Any(field.Name.Contains);
                Formatting.TableDataColumn(Ws.Range[Ws.Cells[lastRow, columnNumber], Ws.Cells[lastRow, columnNumber]], isWide);
            }
        }

        public void UpdateRow(TMessageClass newMessage, TMessageClass oldMessage)
        {
            var row = GetRow(GetUniqueKey(newMessage));
            if (row == 0)
            {
                AppendRow(newMessage);
                return;
            }

            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var field = Fields[fieldIndex];
                var newValue = GetValue(newMessage, field.Accessor);
                var oldValue = GetValue(oldMessage, field.Accessor);
                if (oldValue == newValue) continue;

                var dataCell = Ws.Cells[row, StartColumn + fieldIndex];
                Utilities.AssignCellValue(newMessage, field, newValue, dataCell);
            }

        }

        private int GetRow(object uniqueKey)
        {
            var uniqueKeyString = uniqueKey.ToString();
            var idColumn = 1;
            Range idColumnNameCell = Ws.Cells[HeaderRow, idColumn];
            var uniqueKeyName = Utilities.FormatFieldName(_uniqueKeyName);
            while (idColumnNameCell.Value2 == null || idColumnNameCell.Value2.ToString() != uniqueKeyName)
            {
                idColumn++;
                idColumnNameCell = Ws.Cells[HeaderRow, idColumn];
            }

            var rowNumber = HeaderRow;
            var uniqueKeyCellString = UniqueKeyCellString(Ws.Cells[rowNumber, idColumn]);
            while (uniqueKeyCellString != uniqueKeyString)
            {
                rowNumber++;
                if (rowNumber > 100)
                {
                    return 0;
                }
                uniqueKeyCellString = UniqueKeyCellString(Ws.Cells[rowNumber, idColumn]);
            }
            return rowNumber;
        }

        private static string UniqueKeyCellString(Range uniqueKeyCell)
        {
            if (uniqueKeyCell.Value2 == null) return string.Empty;
            var uniqueKeyCellString = uniqueKeyCell.Value2.ToString();
            return uniqueKeyCellString;
        }

        private int GetLastRow()
        {
            var lastRow = HeaderRow + 1;
            Range dataCell = Ws.Cells[lastRow, StartColumn];
            while (dataCell.Value2 != null && !string.IsNullOrWhiteSpace(dataCell.Value2.ToString()))
            {
                lastRow++;
                dataCell = Ws.Cells[lastRow, StartColumn];
            }
            return lastRow;
        }

        public void Clear()
        {
            var data = Ws.Range[Ws.Cells[HeaderRow + 1, StartColumn], Ws.Cells[GetLastRow(), EndColumn]];
            data.ClearContents();
            DisplayData = new Dictionary<object, TMessageClass>();
            DataList = new List<TMessageClass>();
            DisplayDataList = new List<TMessageClass>();
        }
    }
}
