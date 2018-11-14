using System;
using System.Collections.Generic;
using System.Drawing;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Button = Microsoft.Office.Tools.Excel.Controls.Button;

namespace LightningPower
{
    public class ConnectSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;
        public VerticalTableSheet<GetInfoResponse> GetInfoSheet;


        private readonly int _startRow = 2;
        private readonly int _startColumn = 2;
        private int _dataRow;   
        private readonly Dictionary<string, string> _cellCache;
        private Range _errorData;

        public ConnectSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;
            _cellCache = new Dictionary<string, string>();
            Ws.Change += WsOnChange;

            GetInfoSheet = new VerticalTableSheet<GetInfoResponse>(Ws, LApp, GetInfoResponse.Descriptor);
            GetInfoSheet.SetupVerticalTable("LND Node Info", 18);

            _errorData = Ws.Cells[12, 3];
        }

        private void WsOnChange(Range target)
        {
            string address = target.Address[XlReferenceStyle.xlR1C1];
            if (!_cellCache.TryGetValue(address, out var label)) return;
            var type = LApp.LndClient.Config.GetType();
            var field = type.GetField(label);
            field?.SetValue(LApp.LndClient.Config, target.Value2);
        }

        private void PopulateRow(string label, object value, bool blackOut = false)
        {
            Range labelCell = Ws.Cells[_dataRow, _startColumn];
            labelCell.Value2 = label;
            labelCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
            labelCell.HorizontalAlignment = XlHAlign.xlHAlignRight;
            Formatting.VerticalTableHeaderColumn(labelCell);

            Range dataCell = Ws.Cells[_dataRow, _startColumn + 1];
            Formatting.VerticalTableDataColumn(dataCell);
            Formatting.VerticalTableRow(Ws.Range[labelCell, dataCell], _dataRow);

            if (blackOut)
            {
                dataCell.WrapText = true;
                dataCell.Interior.Color = Color.Gray;
                dataCell.Font.Color = Color.Gray;
            }
            dataCell.NumberFormat = "@";
            dataCell.Value2 = value;
            _cellCache[dataCell.Address[XlReferenceStyle.xlR1C1]] = label;
            _dataRow++;
        }

        public void PopulateConfig()
        {
            _dataRow = _startRow;
            var conf = LApp.LndClient.Config;

            PopulateRow("Network", conf.Network);
            PopulateRow("Host", conf.Host);
            PopulateRow("Port", conf.Port);

            PopulateRow("MacaroonString", conf.MacaroonString, true);
            PopulateRow("MacaroonPath", conf.MacaroonPath);

            PopulateRow("CaCertString", conf.CaCertString, true);
            PopulateRow("CaCertPath", conf.CaCertPath);

            PopulateRow("WalletPassword", conf.WalletPassword, true);
            PopulateRow("BitcoindRpcUser", conf.BitcoindRpcUser, true);
            PopulateRow("BitcoindRpcPassword", conf.BitcoindRpcPassword, true);

            Formatting.VerticalTable(Ws.Range[Ws.Cells[_startRow, _startColumn],
            Ws.Cells[_startRow + 9, _startColumn + 1]]);

            FormatDimensions();
            Button connectButton = Utilities.CreateButton("connect", Ws, Ws.Cells[14, _startColumn], "Connect");
            connectButton.Click += ConnectButtonOnClick;
        }

        public void FormatDimensions()
        {
            Ws.Range["B1"].ColumnWidth = 20;
            Ws.Range["C1"].ColumnWidth = 75;
            Ws.Range["B1:C50"].RowHeight = 14.3;
        }

        private void ConnectButtonOnClick(object sender, EventArgs e)
        {
            //LApp.Connect();
            try
            {
                LApp.Connect();
            }
            catch (RpcException rpcException)
            {
                DisplayRpcError(rpcException);
            }
        }

        public void DisplayRpcError(RpcException exception)
        {
            DisplayError("Connect error", exception.Status.Detail);
        }

        public void DisplayError(string errorType, string errorMessage)
        {
            _errorData.Value2 = $"{errorType}: {errorMessage}";
            Formatting.ActivateErrorCell(_errorData);
            GetInfoSheet.Clear();
        }
    }
}