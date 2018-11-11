using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Grpc.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace LightningPower
{
    public class Utilities
    {
        public static void DisplayError(Range errorData, string errorType, string errorMessage)
        {
            errorData.Value2 = $"{errorType}: {errorMessage}";
            Formatting.ActivateErrorCell(errorData);
        }


        public static void ClearErrorData(Range errorData)
        {
            errorData.Value2 = "";
            Formatting.DeactivateErrorCell(errorData);
        }

        public static void MarkAsLoadingTable(Worksheet ws)
        {
            ws.Cells[1, 2].Value2 = "Loading...";
        }

        public static void RemoveLoadingMark(Worksheet ws)
        {
            try
            {
                ws.Cells[1, 2].Value2 = "";

            }
#pragma warning disable 168
            catch (Exception e)
#pragma warning restore 168
            {
                // ignored
            }
        }

        public static string FormatFieldName(string fieldName)
        {
            return Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(fieldName.Replace("_", " "));
        }

        public static void EnableButton(Worksheet ws, string buttonName, bool enable)
        {
            var worksheet = Globals.Factory.GetVstoObject(ws);
            foreach (Control control in worksheet.Controls)
            {
                if (control.Name == buttonName)
                {
                    control.Enabled = enable;
                }
            }
        }

        public static Microsoft.Office.Tools.Excel.Controls.Button CreateButton(string buttonName, Worksheet ws, Range selection, string buttonText)
        {
            var worksheet = Globals.Factory.GetVstoObject(ws);
            var button = worksheet.Controls.AddButton(selection, buttonName);
            button.Text = buttonText;
            button.Placement = XlPlacement.xlFreeFloating;
            return button;
        }

        public static string GetValue<TMessageClass>(FieldDescriptor field, TMessageClass newMessage) where TMessageClass : IMessage
        {
            var value = "";

            if (field.IsRepeated && field.FieldType == FieldType.String)
            {
                var items = (RepeatedField<string>)field.Accessor.GetValue(newMessage);
                for (var i = 0; i < items.Count; i++)
                {
                    value += items[i];
                    if (i < items.Count - 1)
                    {
                        value += ",\n";
                    }
                }
            }
            else if (field.IsRepeated && field.FieldType == FieldType.Message)
            {
                var enumerable = field.Accessor.GetValue(newMessage) as IEnumerable;
                var items = (enumerable ?? throw new InvalidOperationException()).Cast<object>().ToList();
                return $"{items.Count} {field.Name}";
            }
            else
            {
                return field.Accessor.GetValue(newMessage).ToString();
            }

            return value;
        }

        public static DateTime UnixTimeStampToDateTime(long unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            var dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddSeconds(unixTimeStamp).ToLocalTime();
            return dtDateTime;
        }

        public static void AssignCellValue<TMessageClass>(TMessageClass newMessage, FieldDescriptor field, string newValue, dynamic dataCell) where TMessageClass : IMessage
        {
            var dateFields = new List<string>{"time_stamp", "creation_date", "best_header_timestamp"};
            var isDate = dateFields.Any(field.Name.Contains);
            if (isDate)
            {
                dataCell.NumberFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                var value = UnixTimeStampToDateTime(long.Parse(newValue));
                dataCell.Value2 = value;
            }
            else
            {
                switch (field.FieldType)
                {
                    case FieldType.UInt64:
                        dataCell.NumberFormat = "@";
                        break;
                    case FieldType.Int64:
                        dataCell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)";
                        break;
                }
                dataCell.Value2 = newValue;

            }
        }

        public static void DisplayError(Range errorData, string errorType, RpcException errorMessage)
        {
            DisplayError(errorData, errorType, errorMessage.Status.Detail);
        }
    }
}