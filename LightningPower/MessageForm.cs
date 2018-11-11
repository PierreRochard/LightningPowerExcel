using System;
using System.Collections.Generic;
using Google.Protobuf;
using Google.Protobuf.Reflection;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Button = Microsoft.Office.Tools.Excel.Controls.Button;

namespace LightningPower
{
    public class MessageForm<TRequestMessage, TResponseMessage> where TRequestMessage: IMessage, new()
        where TResponseMessage: IMessage
    {
        public Worksheet Ws;
        private readonly AsyncLightningApp _lApp;
        private readonly Func<TRequestMessage, TResponseMessage> _query;
        public int StartRow;
        public int StartColumn;
        public int EndRow;
        public int EndColumn;
        public IList<FieldDescriptor> Fields;
        public Range ErrorData;
        private readonly int _dataStartRow;
        private readonly Dictionary<string, int> _fieldToRow;
        
        public MessageForm(Worksheet ws, AsyncLightningApp lApp, Func<TRequestMessage, TResponseMessage> query, MessageDescriptor descriptor, string title, int startRow = 2,
            int startColumn = 2)
        {
            Ws = ws;
            Ws.Activate();
            _lApp = lApp;
            _query = query;
            _fieldToRow = new Dictionary<string, int>();
            Fields = descriptor.Fields.InDeclarationOrder();
            StartRow = startRow;
            StartColumn = startColumn;
            EndColumn = StartColumn + 1;

            var titleCell = ws.Cells[StartRow, StartColumn];
            titleCell.Font.Italic = true;
            titleCell.Value2 = title;

            _dataStartRow = StartRow + 1;
            var endDataRow = startRow + Fields.Count;

            var form = ws.Range[ws.Cells[_dataStartRow, StartColumn], ws.Cells[endDataRow, EndColumn]];
            Formatting.VerticalTable(form);

            var header = ws.Range[ws.Cells[_dataStartRow, StartColumn], ws.Cells[endDataRow, StartColumn]];
            Formatting.VerticalTableHeaderColumn(header);

            var data = ws.Range[ws.Cells[_dataStartRow, EndColumn], ws.Cells[endDataRow, EndColumn]];
            Formatting.VerticalTableDataColumn(data);

            var rowNumber = _dataStartRow;
            foreach (var field in Fields)
            {
                var headerCell = ws.Cells[rowNumber, StartColumn];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;
                var rowRange = ws.Range[ws.Cells[rowNumber, StartColumn], ws.Cells[rowNumber, EndColumn]];
                Formatting.VerticalTableRow(rowRange, rowNumber);
                _fieldToRow.Add(field.Name, rowNumber);
                rowNumber++;
            }
            
            var submitButtonRow = rowNumber + 2;
            Button submitButton = Utilities.CreateButton("submit" + descriptor.Name, ws, ws.Cells[submitButtonRow, StartColumn], "Submit");
            submitButton.Click += SubmitButtonOnClick;
            ErrorData = ws.Cells[submitButtonRow, StartColumn + 2];
            ErrorData.WrapText = false;
            ErrorData.RowHeight = 14.3;

            EndRow = submitButtonRow + 1;

            titleCell.Columns.AutoFit();
        }

        public void ClearErrorData()
        {
            Utilities.ClearErrorData(ErrorData);
            ErrorData.Columns.AutoFit();
        }

        private void SubmitButtonOnClick(object sender, EventArgs e)
        {
            ClearErrorData();
            if (typeof(TRequestMessage) == typeof(ConnectPeerRequest))
            {
                var fullAddress = (string) Ws.Cells[_fieldToRow["addr"], EndColumn].Value2;
                if (fullAddress == null) return;
                var addressParts = fullAddress.Split('@');

                string pubkey;
                string host;
                switch (addressParts.Length)
                {
                    case 0:
                        return;
                    case 2:
                        pubkey = addressParts[0];
                        host = addressParts[1];
                        break;
                    default:
                        Utilities.DisplayError(ErrorData, "Error", "Invalid address, must be pubkey@ip:host");
                        return;
                }

                var permanent = Ws.Cells[_fieldToRow["perm"], EndColumn].Value2;
                bool perm = permanent == null || (bool) permanent;

                var address = new LightningAddress { Host = host, Pubkey = pubkey };
                var request = new ConnectPeerRequest { Addr = address, Perm = perm };
                try
                {
                    _lApp.LndClient.ConnectPeer(request);
                    _lApp.Refresh(SheetNames.Peers);
                    ClearForm();
                }
                catch (RpcException rpcException)
                {
                    DisplayError(rpcException);
                }
            }
            else if (typeof(TRequestMessage) == typeof(SendCoinsRequest))
            {
                var request = new SendCoinsRequest
                {
                    Addr = Ws.Cells[_fieldToRow["addr"], EndColumn].Value2,
                    Amount = (long)Ws.Cells[_fieldToRow["amount"], EndColumn].Value2
                };
                var satPerByte = Ws.Cells[_fieldToRow["sat_per_byte"], EndColumn].Value2;
                if (satPerByte == null) satPerByte = 0;
                if (satPerByte > 0) request.SatPerByte = satPerByte;

                var targetConf = Ws.Cells[_fieldToRow["target_conf"], EndColumn].Value2;
                if (targetConf == null) targetConf = 0;
                if (targetConf > 0) request.TargetConf = targetConf;

                try
                {
                    _lApp.LndClient.SendCoins(request);
                    _lApp.Refresh(SheetNames.Transactions);
                    ClearForm();
                }
                catch (RpcException rpcException)
                {
                    DisplayError(rpcException);
                }
            }
            else if (typeof(TRequestMessage) == typeof(OpenChannelRequest))
            {
                var localFundingAmount = long.Parse(GetValue("local_funding_amount"));
                var minConfs = int.Parse(GetValue("min_confs"));
                var minHtlcMsat = long.Parse(GetValue("min_htlc_msat"));
                var nodePubKeyString = GetValue("node_pubkey");
                var isPrivate = true;
                if (bool.TryParse(GetValue("private"), out var result)) isPrivate = result;
                var pushSat = long.Parse(GetValue("push_sat"));
                var remoteCsvDelay = uint.Parse(GetValue("remote_csv_delay"));
                var satPerByte = long.Parse(GetValue("sat_per_byte"));
                var targetConf = int.Parse(GetValue("target_conf"));
                var request = new OpenChannelRequest
                {
                    LocalFundingAmount = localFundingAmount,
                    MinConfs = minConfs,
                    MinHtlcMsat = minHtlcMsat,
                    NodePubkeyString = nodePubKeyString,
                    Private = isPrivate,
                    PushSat = pushSat
                };
                if (remoteCsvDelay > 0) request.RemoteCsvDelay = remoteCsvDelay;
                if (satPerByte > 0) request.SatPerByte = satPerByte;
                if (targetConf > 0) request.TargetConf = targetConf;

                try
                {
                    _lApp.LndClient.OpenChannel(request);
                    _lApp.Refresh(SheetNames.Channels);
                    ClearForm();
                }
                catch (RpcException rpcException)
                {
                    DisplayError(rpcException);
                }
            }
            else
            {
                var request = new TRequestMessage();
                var rowNumber = _dataStartRow;
                foreach (var field in Fields)
                {
                    Range dataCell = Ws.Cells[rowNumber, EndColumn];
                    var value = dataCell.Value2;
                    if (!string.IsNullOrWhiteSpace(value?.ToString()))
                    {
                        field.Accessor.SetValue(request, dataCell.Value2);
                    }
                }

                _query(request);
            }
        }

        private string GetValue(string name)
        {
            var value = Ws.Cells[_fieldToRow[name], EndColumn].Value2;
            var adjValue = value ?? "0";
            return adjValue.ToString();
        }

        private void ClearForm()
        {
            var rowNumber = _dataStartRow;
            foreach (var _ in Fields)
            {
                var dataCell = Ws.Cells[rowNumber, EndColumn];
                dataCell.Value2 = "";
                rowNumber++;
            }
        }

        public void DisplayError(RpcException e)
        {
            Utilities.DisplayError(ErrorData, "Error", e);
            ErrorData.Columns.AutoFit();
        }
    }
}