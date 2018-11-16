using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public class MnemonicSheet
    {
        public Worksheet Ws;
        public int StartRow = 2;
        public int StartColumn = 2;

        public MnemonicSheet(Worksheet ws)
        {
            Ws = ws;
        }

        public void Populate(List<string> mnemonic)
        {
            var titleCell = Ws.Cells[StartRow, StartColumn];
            titleCell.Value2 = "Backup seed mnemonic, print for your records!";
            for (int i = 0; i < mnemonic.Count; i++)
            {
                var labelCell = Ws.Cells[StartRow + 1 + i, StartColumn];
                var dataCell = Ws.Cells[StartRow + 1 + i, StartColumn + 1];
                labelCell.Value2 = i + 1;
                dataCell.Value2 = mnemonic[i];
            }

            Ws.Activate();
            MessageBox.Show(@"Please print this sheet to backup your private key's mnemonic seed!");
        }
    }
}