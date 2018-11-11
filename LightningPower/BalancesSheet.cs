using System;
using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public class BalancesSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;
        public VerticalTableSheet<ChannelBalanceResponse> ChannelBalanceSheet;
        public VerticalTableSheet<WalletBalanceResponse> WalletBalanceSheet;

        public BalancesSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;

            WalletBalanceSheet = new VerticalTableSheet<WalletBalanceResponse>(Ws, LApp, WalletBalanceResponse.Descriptor);
            WalletBalanceSheet.SetupVerticalTable("Wallet Balance");

            ChannelBalanceSheet = new VerticalTableSheet<ChannelBalanceResponse>(Ws, LApp, ChannelBalanceResponse.Descriptor);
            ChannelBalanceSheet.SetupVerticalTable("Channel Balance", startColumn: 5);

            Ws.Columns.AutoFit();

            //Ws.Names.Add(field.Name, dataCell);
        }

        public void Update(Tuple<WalletBalanceResponse, ChannelBalanceResponse> result)
        {
            ChannelBalanceSheet.Update(result.Item2);
            WalletBalanceSheet.Update(result.Item1);
            Utilities.RemoveLoadingMark(Ws);
            Ws.Columns.AutoFit();
        }
    }
}