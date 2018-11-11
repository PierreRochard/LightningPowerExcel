using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace LightningPower
{
    public class NodeSheet
    {
        public Worksheet Ws;
        public bool IsProcessOurs = false;

        public NodeSheet(Worksheet ws)
        {
            Ws = ws;
        }

        public void StartLocalNode(LndClientConfiguration conf)
        {
            var bw = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            bw.DoWork += (sender, args) => RunLnd(sender, args, conf.Network, conf.BitcoindRpcUser, conf.BitcoindRpcPassword, conf.Autopilot);
            bw.ProgressChanged += BwRunLndOnProgressChanged;
            bw.RunWorkerAsync();
        }

        private void BwRunLndOnProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            WriteToLog(e.UserState);
        }

        // ReSharper disable once UnusedParameter.Local
        private void RunLnd(object sender, DoWorkEventArgs e, string network, string rpcUser, string rpcPassword, bool autopilot)
        {
            if (IsProcessOurs) return;

            var processes = Process.GetProcessesByName("tempfileLND");
            foreach (var t in processes)
            {
                t.Kill();
            }

            var lndProcesses = Process.GetProcessesByName("lnd");
            if (lndProcesses.Length > 0)
            {
                WriteToLog("LND is already running, not spawning a process and thus unable to redirect log output to this tab.");
                return;
            }

            const string exeName = "tempfileLND.exe";
            var path = Path.Combine(Path.GetTempPath(), exeName);
            try
            {
                File.WriteAllBytes(path, Properties.Resources.lnd);
            }
#pragma warning disable 168
            catch (IOException exception)
#pragma warning restore 168
            {
                return;
            }

            var cmdArgs = "--bitcoin.active " +
                          $"--bitcoin.{network} " +
                          "--bitcoin.node=bitcoind " +
                          "--bitcoind.rpchost=127.0.0.1 " +
                          $"--bitcoind.rpcuser={rpcUser} " +
                          $"--bitcoind.rpcpass={rpcPassword} " +
                          "--bitcoind.zmqpubrawblock=tcp://127.0.0.1:18501 " +
                          "--bitcoind.zmqpubrawtx=tcp://127.0.0.1:18502" +
                          "--debuglevel=info ";
            if (autopilot)
            {
                cmdArgs += "--autopilot.active " +
                           "--autopilot.maxchannels=10 " +
                           "--autopilot.allocation=1 " +
                           "--autopilot.minchansize=600000 " +
                           "--autopilot.private ";
            }

            var nodeProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = path,
                    Arguments = cmdArgs,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };
            nodeProcess.Start();
            IsProcessOurs = true;
            nodeProcess.EnableRaisingEvents = true;
            nodeProcess.OutputDataReceived += (o, args) =>
                NodeProcessOutputDataReceived(o, args, (BackgroundWorker) sender);
            nodeProcess.ErrorDataReceived += (o, args) =>
                NodeProcessOutputDataReceived(o, args, (BackgroundWorker) sender);
            nodeProcess.BeginOutputReadLine();
            nodeProcess.BeginErrorReadLine();
            nodeProcess.WaitForExit();
        }

        private void WriteToLog(object logMessage)
        {
            var line = (Range)Ws.Rows[1];
            try
            {
                line.Insert(XlInsertShiftDirection.xlShiftDown);
                var cell = Ws.Cells[1, 1];
                cell.Value2 = logMessage;
            }
            catch (Exception)
            {
                // ignored
            }

        }

        // ReSharper disable once UnusedParameter.Local
        private void NodeProcessOutputDataReceived(object sender, DataReceivedEventArgs e, BackgroundWorker bw)
        {
            bw.ReportProgress(0, e.Data);
        }
        
        
    }
}