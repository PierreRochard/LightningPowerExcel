using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public class TransactionsSheet
    {
        public MessageForm<SendCoinsRequest, SendCoinsResponse> SendCoinsForm;
        public TableSheet<Transaction> TransactionsTable;

        public TransactionsSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            SendCoinsForm = new MessageForm<SendCoinsRequest, SendCoinsResponse>(ws, lApp, lApp.LndClient.SendCoins, SendCoinsRequest.Descriptor, "Send on-chain bitcoins");
            TransactionsTable = new TableSheet<Transaction>(ws, lApp, Transaction.Descriptor, "tx_hash");
            TransactionsTable.SetupTable("Transactions", startRow:SendCoinsForm.EndRow + 2);
        }
    }
}