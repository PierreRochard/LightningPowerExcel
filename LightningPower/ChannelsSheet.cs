using System;
using System.Collections.Generic;
using System.Linq;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using static Lnrpc.PendingChannelsResponse.Types;
using Channel = Lnrpc.Channel;

namespace LightningPower
{
    public class ChannelsSheet
    {
        public MessageForm<OpenChannelRequest, ChannelPoint> OpenChannelForm;
        public TableSheet<PendingOpenChannel> PendingOpenChannelsTable;
        public TableSheet<Channel> OpenChannelsTable;
        public TableSheet<ClosedChannel> PendingClosingChannelsTable;
        public TableSheet<ForceClosedChannel> PendingForceClosingChannelsTable;
        public TableSheet<WaitingCloseChannel> WaitingCloseChannelsTable;
        public TableSheet<ChannelCloseSummary> ClosedChannelsTable;

        public int StartColumn = 2;
        public int EndColumn;
        public int EndRow;
        public AsyncLightningApp LApp;

        public ChannelsSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            LApp = lApp;
            OpenChannelForm = new MessageForm<OpenChannelRequest, ChannelPoint>(ws, lApp, lApp.LndClient.OpenChannel,
                OpenChannelRequest.Descriptor, "Open a new channel");

            PendingOpenChannelsTable = new TableSheet<PendingOpenChannel>(ws, lApp, PendingOpenChannel.Descriptor, "channel_point", true);
            PendingOpenChannelsTable.SetupTable("Pending open", 5, OpenChannelForm.EndRow + 2);

            OpenChannelsTable = new TableSheet<Channel>(ws, lApp, Channel.Descriptor, "chan_id");
            OpenChannelsTable.SetupTable("Open", 10, PendingOpenChannelsTable.EndRow + 2);

            PendingClosingChannelsTable = new TableSheet<ClosedChannel>(ws, lApp, ClosedChannel.Descriptor, "channel_point", true);
            PendingClosingChannelsTable.SetupTable("Pending closing", 5, OpenChannelsTable.EndRow+2, StartColumn);

            PendingForceClosingChannelsTable = new TableSheet<ForceClosedChannel>(ws, lApp, ForceClosedChannel.Descriptor, "channel_point", true);
            PendingForceClosingChannelsTable.SetupTable("Pending force closing", 5, PendingClosingChannelsTable.EndRow+2, StartColumn);

            WaitingCloseChannelsTable = new TableSheet<WaitingCloseChannel>(ws, lApp, WaitingCloseChannel.Descriptor, "channel_point", true);
            WaitingCloseChannelsTable.SetupTable("Waiting for closing transaction to confirm", 5, PendingForceClosingChannelsTable.EndRow + 2, StartColumn);

            ClosedChannelsTable = new TableSheet<ChannelCloseSummary>(ws, lApp, ChannelCloseSummary.Descriptor, "chan_id");
            ClosedChannelsTable.SetupTable("Closed", 5, WaitingCloseChannelsTable.EndRow + 2);
            
            EndRow = WaitingCloseChannelsTable.EndRow;
            EndColumn = new List<int>
            {
                PendingOpenChannelsTable.EndColumn,
                PendingClosingChannelsTable.EndColumn,
                PendingForceClosingChannelsTable.EndColumn,
                WaitingCloseChannelsTable.EndColumn,
                ClosedChannelsTable.EndColumn,
                OpenChannelsTable.EndColumn
            }.Max();

            ws.Change += WsOnChange;
        }

        public void Update(Tuple<ListChannelsResponse, PendingChannelsResponse, ClosedChannelsResponse> r)
        {
            OpenChannelsTable.Update(r.Item1.Channels);
            var pendingChannels = r.Item2;
            PendingOpenChannelsTable.Update(pendingChannels.PendingOpenChannels);
            PendingClosingChannelsTable.Update(pendingChannels.PendingClosingChannels);
            PendingForceClosingChannelsTable.Update(pendingChannels.PendingForceClosingChannels);
            WaitingCloseChannelsTable.Update(pendingChannels.WaitingCloseChannels);
            ClosedChannelsTable.Update(r.Item3.Channels);
        }

        private void WsOnChange(Range target)
        {
            if (target.Row < OpenChannelsTable.StartRow || target.Row > OpenChannelsTable.EndRow ||
                (target.Value2?.ToString() != "" && target.Value2 != null)) return;
            OpenChannelForm.ClearErrorData();
            var channel = OpenChannelsTable.DataList[target.Row - OpenChannelsTable.DataStartRow];
            try
            {
                LApp.LndClient.CloseChannel(channel.ChannelPoint, false);
                OpenChannelsTable.RemoveRow(target.Row);
                LApp.Refresh(SheetNames.Channels);
            }
            catch (RpcException e)
            {
                if (e.Status.Detail.Contains("peer is offline"))
                {
                    try
                    {
                        LApp.LndClient.CloseChannel(channel.ChannelPoint, true);
                        OpenChannelsTable.RemoveRow(target.Row);
                        LApp.Refresh(SheetNames.Channels);
                    }
                    catch (RpcException e2)
                    {
                        OpenChannelForm.DisplayError(e2);
                    }
                }
                else
                {
                    OpenChannelForm.DisplayError(e);
                }
            }
        }
    }
}