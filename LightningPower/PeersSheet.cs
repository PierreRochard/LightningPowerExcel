using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public class PeersSheet
    {
        public AsyncLightningApp LApp;

        public MessageForm<ConnectPeerRequest, ConnectPeerResponse> PeersForm;
        public TableSheet<Peer> PeersTable;
        
        public PeersSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            LApp = lApp;

            PeersForm = new MessageForm<ConnectPeerRequest, ConnectPeerResponse>(ws, LApp, LApp.LndClient.ConnectPeer, ConnectPeerRequest.Descriptor, "Connect to a peer");
            PeersTable = new TableSheet<Peer>(ws, LApp, Peer.Descriptor, "pub_key");
            PeersTable.SetupTable("Peers", startRow: PeersForm.EndRow + 2);

            ws.Change += WsOnChange;
        }

        private void WsOnChange(Range target)
        {
            if (target.Row < PeersTable.StartRow || target.Row > PeersTable.EndRow ||
                (target.Value2?.ToString() != "" && target.Value2 != null)) return;
            PeersForm.ClearErrorData();
            try
            {
                var peer = PeersTable.DataList[target.Row - PeersTable.DataStartRow];
                LApp.LndClient.DisconnectPeer(peer.PubKey);
                PeersTable.RemoveRow(target.Row);
                LApp.Refresh(SheetNames.Peers);
            }
            catch (RpcException e)
            {
                PeersForm.DisplayError(e);
            }
        }
    }
}