using System;
using System.Collections.Generic;
using System.Drawing;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Button = Microsoft.Office.Tools.Excel.Controls.Button;

namespace LightningPower
{
    public class SendPaymentSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        public VerticalTableSheet<PayReq> PaymentRequestTable;
        public VerticalTableSheet<Route> RouteTakenTable;
        public TableSheet<Hop> HopTable;
        public TableSheet<Route> PotentialRoutesTable;

        private const int MaxRoutes = 10;

        private Range _payReqLabelCell;
        private Range _payReqInputCell;
        private Range _payReqRange;

        private Range _errorData;

        private Range _sendStatusRange;

        private Range _paymentPreimageCell;
        private Range _paymentPreimageLabel;

        public int StartColumn = 2;
        public int StartRow = 2;

        private const int PayReqColumnWidth = 15;

        public SendPaymentSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;
        }
        
        public void InitializePaymentRequest()
        {
            Ws.Activate();
            _payReqLabelCell = Ws.Cells[StartRow, StartColumn];
            _payReqLabelCell.Value2 = "Payment request:";
            _payReqLabelCell.Font.Bold = true;
            _payReqLabelCell.Columns.AutoFit();

            _payReqInputCell = Ws.Cells[StartRow, StartColumn + 1];
            _payReqInputCell.Interior.Color = Color.AliceBlue;
            Formatting.WideTableColumn(_payReqInputCell);

            _payReqRange = Ws.Range[_payReqLabelCell, _payReqInputCell];
            Formatting.UnderlineBorder(_payReqRange);

            Ws.Change += WsOnChangeParsePayReq;

            PaymentRequestTable = new VerticalTableSheet<PayReq>(Ws, LApp, PayReq.Descriptor);
            PaymentRequestTable.SetupVerticalTable("Decoded Payment Request", StartRow + 2);

            PotentialRoutesTable = new TableSheet<Route>(Ws, LApp, Route.Descriptor, "hops");
            PotentialRoutesTable.SetupTable("Potential Routes", MaxRoutes, StartRow=PaymentRequestTable.EndRow + 2);
            
            var sendPaymentButtonRow = PotentialRoutesTable.EndRow + 4;
            Button sendPaymentButton = Utilities.CreateButton("sendPayment", Ws, Ws.Cells[sendPaymentButtonRow, StartColumn], "Send Payment");
            sendPaymentButton.Click += SendPaymentButtonOnClick;
            _errorData = Ws.Cells[sendPaymentButtonRow + 3, StartColumn + 1];

            _sendStatusRange = Ws.Cells[sendPaymentButtonRow + 3, StartColumn];
            _sendStatusRange.Font.Italic = true;

            Button clearPaymentInfoButton = Utilities.CreateButton("clearPaymentInfo", Ws, Ws.Cells[sendPaymentButtonRow + 3, StartColumn], "Clear");
            clearPaymentInfoButton.Click += ClearPaymentInfoButtonOnClick;
            
            var paymentResponseDataStartRow = sendPaymentButtonRow + 5;
            _paymentPreimageLabel = Ws.Cells[paymentResponseDataStartRow, StartColumn];
            _paymentPreimageLabel.Value2 = "Proof of Payment";
            _paymentPreimageLabel.Font.Italic = true;
            Formatting.WideTableColumn(_paymentPreimageLabel);

            _paymentPreimageCell = Ws.Cells[paymentResponseDataStartRow, StartColumn + 1];
            _paymentPreimageCell.Interior.Color = Color.PaleGreen;
            _paymentPreimageCell.RowHeight = 14.3;
            _paymentPreimageCell.WrapText = true;
            
            RouteTakenTable = new VerticalTableSheet<Route>(Ws, LApp, Route.Descriptor, new List<string> { "hops" });
            RouteTakenTable.SetupVerticalTable("Payment Summary", paymentResponseDataStartRow + 3);

            HopTable = new TableSheet<Hop>(Ws, LApp, Hop.Descriptor, "chan_id");
            HopTable.SetupTable("Route", 4, RouteTakenTable.EndRow + 2);

            _payReqInputCell.Columns.ColumnWidth = PayReqColumnWidth;
            Utilities.RemoveLoadingMark(Ws);
        }


        private void ClearPaymentInfoButtonOnClick(object sender, EventArgs e)
        {
            ClearPayReq();
            PaymentRequestTable.Clear();
            PotentialRoutesTable.Clear();
            ClearSendStatus();
            Utilities.ClearErrorData(_errorData);
            ClearSendPaymentResponseData();
        }

        private void ClearPayReq()
        {
            _payReqInputCell.Value2 = "";
        }


        private void ClearSendStatus()
        {
            _sendStatusRange.Value2 = "";
        }

        private void ClearSendPaymentResponseData()
        {
            _paymentPreimageCell.Value2 = "";
            RouteTakenTable.Clear();
            HopTable.Clear();
        }

        private void WsOnChangeParsePayReq(Range target)
        {
            
            if (target.Address != "$C$2")
            {
                return;
            }

            string payReq = target.Value2;
            if (string.IsNullOrWhiteSpace(payReq))
            {
                return;
            }

            PayReq response;
            try
            {
                response = LApp.DecodePaymentRequest(payReq);
                PaymentRequestTable.Clear();
                PotentialRoutesTable.Clear();
                ClearSendStatus();
                Utilities.ClearErrorData(_errorData);
                ClearSendPaymentResponseData();

            }
            catch (RpcException e)
            {
                Utilities.DisplayError(_errorData, "Parsing error", e);
                return;
            }
            PaymentRequestTable.Update(response);
            Utilities.ClearErrorData(_errorData);

            try
            {
               var r = LApp.QueryRoutes(response, MaxRoutes);
               PotentialRoutesTable.Update(r.Routes);
            }
            catch (RpcException e)
            {
                Utilities.DisplayError(_errorData, "Query route error", e);
                return;
            }

            _payReqInputCell.Columns.ColumnWidth = PayReqColumnWidth;
        }

        private void SendPaymentButtonOnClick(object sender, EventArgs e)
        { 
            // Disable the Send Payment button so that it's not clicked twice
            Utilities.EnableButton(Ws, "sendPayment", false);

            string payReq = _payReqInputCell.Value2;
            if (string.IsNullOrWhiteSpace(payReq))
            {
                return;
            }

            try
            {
                LApp.SendPayment(payReq);
            }
            catch (RpcException rpcException)
            {
                Utilities.DisplayError(_errorData, "Payment error", rpcException);
            }
        }

        public void MarkSendingPayment()
        {
            ClearSendPaymentResponseData();
            Utilities.ClearErrorData(_errorData);
            // Indicate payment is being sent below send button
            _sendStatusRange.Value2 = "Sending payment...";
        }

        public void PopulateSendPaymentError(RpcException exception)
        {
            Utilities.DisplayError(_errorData, "Payment error", exception);
            ClearSendStatus();
        }

        public void PopulateSendPaymentResponse(SendResponse response)
        {
            ClearSendStatus();
            if (response.PaymentError == "")
            {
                _paymentPreimageCell.Value2 = BitConverter.ToString(response.PaymentPreimage.ToByteArray()).Replace("-", "").ToLower();
                _paymentPreimageCell.RowHeight = 14.3;

                RouteTakenTable.Populate(response.PaymentRoute);
                HopTable.Update(response.PaymentRoute.Hops);
                _payReqInputCell.Columns.ColumnWidth = PayReqColumnWidth;
            }
            else
            {
                Utilities.DisplayError(_errorData, "Payment error", response.PaymentError);
            }
            Utilities.EnableButton(Ws, "sendPayment", true);
        }

        public void UpdateSendPaymentProgress(int progress)
        {
            // Indicate payment is being sent below send button
            _sendStatusRange.Value2 = $"Sending payment...{progress}%";
        }


    }
}