using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public partial class ThisAddIn
    {

        public AsyncLightningApp LApp;
        public Workbook Wb;

        public MnemonicSheet MnemonicSheet;
        public ConnectSheet ConnectSheet;
        public PeersSheet PeersSheet;
        public BalancesSheet BalancesSheet;
        public TransactionsSheet TransactionsSheet;
        public ChannelsSheet ChannelsSheet;
        public TableSheet<Payment> PaymentsSheet;
        public SendPaymentSheet SendPaymentSheet;
        public NodeSheet NodesSheet;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WorkbookOpen += ApplicationOnWorkbookOpen;
        }

        private bool IsLndWorkbook()
        {
            try
            {
                if (Application.Sheets[SheetNames.Connect].Cells[1, 1].Value2 == "LightningPower")
                {
                    return true;
                }
            }
            catch (COMException)
            {
                // GetInfo tab doesn't exist, certainly not an LightningPower workbook
            }

            return false;
        }

        // Check to see if the workbook is an LightningPower workbook
        private void ApplicationOnWorkbookOpen(Workbook wb)
        {
            if (IsLndWorkbook())
            {
                SetupWorkbook(wb);
            }
        }

        public void SetupWorkbook(Workbook wb)
        {
            Wb = wb;
            LApp = new AsyncLightningApp(this);

            CreateSheet(SheetNames.Mnemonic);
            MnemonicSheet = new MnemonicSheet(Wb.Sheets[SheetNames.Mnemonic]);

            CreateSheet(SheetNames.Connect);
            ConnectSheet = new ConnectSheet(Wb.Sheets[SheetNames.Connect], LApp);
            ConnectSheet.PopulateConfig();

            CreateSheet(SheetNames.Peers);
            PeersSheet = new PeersSheet(Wb.Sheets[SheetNames.Peers], LApp);

            CreateSheet(SheetNames.Balances);
            BalancesSheet = new BalancesSheet(Wb.Sheets[SheetNames.Balances], LApp);

            CreateSheet(SheetNames.Transactions);
            TransactionsSheet = new TransactionsSheet(Wb.Sheets[SheetNames.Transactions], LApp);

            CreateSheet(SheetNames.Channels);
            ChannelsSheet = new ChannelsSheet(Wb.Sheets[SheetNames.Channels], LApp);

            CreateSheet(SheetNames.Payments);
            PaymentsSheet = new TableSheet<Payment>(Wb.Sheets[SheetNames.Payments], LApp, Payment.Descriptor, "payment_hash");
            PaymentsSheet.SetupTable("Payments");

            CreateSheet(SheetNames.SendPayment);
            SendPaymentSheet = new SendPaymentSheet(Wb.Sheets[SheetNames.SendPayment], LApp);
            SendPaymentSheet.InitializePaymentRequest();

            CreateSheet(SheetNames.NodeLog);
            NodesSheet = new NodeSheet(Wb.Sheets[SheetNames.NodeLog]);

            MarkLightningPowerWorkbook();
            ConnectSheet.Ws.Activate();

            Application.SheetActivate += Workbook_SheetActivate;
        }

        private void CreateSheet(string worksheetName)
        {
            Worksheet oldWs = Wb.ActiveSheet;
            Worksheet ws;
            try
            {
                // ReSharper disable once RedundantAssignment
                ws = Wb.Sheets[worksheetName];
            }
            catch (COMException)
            {
                Globals.ThisAddIn.Wb.Sheets.Add(After: Wb.Sheets[Wb.Sheets.Count]);
                ws = Wb.ActiveSheet;
                ws.Name = worksheetName;
                ws.Range["A:AZ"].Interior.Color = Color.White;
            }
            oldWs.Activate();
        }

        private void Workbook_SheetActivate(object sh)
        {
            if (!IsLndWorkbook())
            {
                return;
            }
            var ws = (Worksheet)sh;
            LApp?.Refresh(ws.Name);
        }

        private void MarkLightningPowerWorkbook()
        {
            Worksheet ws = Wb.Sheets[SheetNames.Connect];
            ws.Cells[1, 1].Value2 = "LightningPower";
            ws.Cells[1, 1].Font.Color = Color.White;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (!NodesSheet.IsProcessOurs) return;
            }
            catch (NullReferenceException)
            {
                return;
            }

            try
            {
                LApp.StopDaemon();
            }
            catch (RpcException)
            {
                NodesSheet.IsProcessOurs = false;
            }
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}