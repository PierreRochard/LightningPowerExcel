using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Grpc.Core;
using Lnrpc;
using Channel = Grpc.Core.Channel;

namespace LightningPower
{
    public class LndClientConfiguration
    {
        private string _macaroonPath;
        private string _macaroonString;
        private string _caCertPath;
        private string _caCertString;

        public string Network = "mainnet";
        public string Host = "localhost";
        public double Port = 10009;
        public string WalletPassword = "test_password";
        public bool Autopilot = false;

        public string BitcoindRpcUser = "test_user";
        public string BitcoindRpcPassword = "test_password";


        public static string LndDataPath
        {
            get
            {
                var localAppData = Environment.GetEnvironmentVariable("LocalAppData");
                string[] lndPaths = { localAppData, "Lnd" };
                var lndPath = Path.Combine(lndPaths);
                return lndPath;
            }
        }

        public string MacaroonString
        {
            get
            {
                if (_macaroonString != null) return _macaroonString;
                try
                {
                    var macaroonBytes = File.ReadAllBytes(MacaroonPath);
                    var macaroonString = BitConverter.ToString(macaroonBytes).Replace("-", "").ToLower();
                    return macaroonString;
                }
                catch (FileNotFoundException)
                {
                    return null;
                }
                catch (DirectoryNotFoundException)
                {
                    return null;
                }
            }
            set => _macaroonString = value;
        }

        public string MacaroonPath
        {
            get
            {
                if (_macaroonPath != null) return _macaroonPath;
                string[] macaroonPaths = { LndDataPath, "data", "chain", "bitcoin", Network, "admin.macaroon" };
                var macaroonPath = Path.Combine(macaroonPaths);
                return macaroonPath;
            }
            set => _macaroonPath = value;
        }

        public SslCredentials SslCredentials
        {
            get
            {
                var ssl = new SslCredentials(CaCertString);
                return ssl;
            }
        }

        public string CaCertPath
        {
            get
            {
                if (_caCertPath != null) return _caCertPath;

                string[] caCertPaths = { LndDataPath, "tls.cert" };
                var caCertPath = Path.Combine(caCertPaths);
                return caCertPath;
            }
            set => _caCertPath = value;
        }

        public string CaCertString
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(_caCertString)) return _caCertString;
                try
                {
                    var caCert = File.ReadAllText(CaCertPath);
                    return caCert;
                }
                catch (FileNotFoundException)
                {
                    return null;
                }
                catch (DirectoryNotFoundException)
                {
                    return null;
                }
            }
            set => _caCertString = value;
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public async Task AsyncAuthInterceptor(AuthInterceptorContext context, Metadata metadata)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            if (!string.IsNullOrWhiteSpace(MacaroonString))
            {
                metadata.Add(new Metadata.Entry("macaroon", MacaroonString));
            }
        }

        public Channel RpcChannel
        {
            get
            {
                var callCredentials = CallCredentials.FromInterceptor(AsyncAuthInterceptor);
                var channelCredentials = ChannelCredentials.Create(SslCredentials, callCredentials);
                var channel = new Channel(Host, (int)Port, channelCredentials);
                return channel;
            }
        }

    }

    public class LndClient
    {
        public LndClientConfiguration Config;

        public LndClient()
        {
            Environment.SetEnvironmentVariable("GRPC_SSL_CIPHER_SUITES", "HIGH+ECDSA", EnvironmentVariableTarget.Process);
            Config = new LndClientConfiguration();
        }

        public List<string> TryUnlockWallet(string password)
        {
            List<string> mnemonic = new List<string>();
            if (Config.MacaroonString == null)
            {
                var seed = GenerateSeed();
                mnemonic = seed.CipherSeedMnemonic.ToList();
                InitWallet(Config.WalletPassword, seed.CipherSeedMnemonic);
                Thread.Sleep(3000);
            }
            try
            {
                // ReSharper disable once UnusedVariable
                var response = UnlockWallet(password);
                Thread.Sleep(3000);
            }
            catch (RpcException e)
            {
                if ("unknown service lnrpc.WalletUnlocker" == e.Status.Detail)
                {
                    // Wallet is already unlocked
                }
                else
                {
                   // throw;
                }
            }

            return mnemonic;
        }

        public Lightning.LightningClient GetLightningClient()
        {
            return new Lightning.LightningClient(Config.RpcChannel);
        }

        public WalletUnlocker.WalletUnlockerClient GetWalletUnlockerClient()
        {
            return new WalletUnlocker.WalletUnlockerClient(Config.RpcChannel);
        }

        public GenSeedResponse GenerateSeed()
        {
            var request = new GenSeedRequest();
            var response = GetWalletUnlockerClient().GenSeed(request);
            return response;
        }

        public void InitWallet(string walletPassword, RepeatedField<string> seed)
        {
            var request = new InitWalletRequest
            {
                WalletPassword = ByteString.CopyFrom(walletPassword, Encoding.UTF8)
            };
            request.CipherSeedMnemonic.Add(seed);
            GetWalletUnlockerClient().InitWalletAsync(request);
        }

        public StopResponse StopDaemon()
        {
            var request = new StopRequest();
            var response = GetLightningClient().StopDaemon(request);
            return response;
        }

        public UnlockWalletResponse UnlockWallet(string password)
        {
            var pw = ByteString.CopyFrom(password, Encoding.UTF8);
            var req = new UnlockWalletRequest { WalletPassword = pw };
            var response = GetWalletUnlockerClient().UnlockWallet(req);
            return response;
        }

        public GetInfoResponse GetInfo()
        {
            var request = new GetInfoRequest();
            var response = GetLightningClient().GetInfo(request);
            return response;
        }

        public ConnectPeerResponse ConnectPeer(ConnectPeerRequest request)
        {
            var response = GetLightningClient().ConnectPeer(request);
            return response;
        }

        public DisconnectPeerResponse DisconnectPeer(string pubkey)
        {
            var request = new DisconnectPeerRequest{PubKey = pubkey};
            var response = GetLightningClient().DisconnectPeer(request);
            return response;
        }

        public ListPeersResponse ListPeers()
        {
            var request = new ListPeersRequest();
            var response = GetLightningClient().ListPeers(request);
            return response;
        }

        public WalletBalanceResponse WalletBalance()
        {
            var request = new WalletBalanceRequest();
            var response = GetLightningClient().WalletBalance(request);
            return response;
        }

        public SendCoinsResponse SendCoins(SendCoinsRequest request)
        {
            var response = GetLightningClient().SendCoins(request);
            return response;
        }

        public TransactionDetails GetTransactions()
        {
            var request = new GetTransactionsRequest();
            var response = GetLightningClient().GetTransactions(request);
            return response;
        }

        public IAsyncStreamReader<Transaction> SubscribeTransactions()
        {
            var request = new GetTransactionsRequest();
            var response = GetLightningClient().SubscribeTransactions(request);
            return response.ResponseStream;
        }

        public NewAddressResponse NewAddress(NewAddressRequest.Types.AddressType addressType = NewAddressRequest.Types.AddressType.WitnessPubkeyHash)
        {
            var request = new NewAddressRequest { Type = addressType };
            var response = GetLightningClient().NewAddress(request);
            return response;
        }

        public SignMessageResponse SignMessage(string message)
        {
            var request = new SignMessageRequest
            {
                Msg = ByteString.CopyFrom(message, Encoding.UTF8)
            };
            var response = GetLightningClient().SignMessage(request);
            return response;
        }

        public VerifyMessageResponse VerifyMessage(string message, string signature)
        {
            var request = new VerifyMessageRequest
            {
                Msg = ByteString.CopyFrom(message, Encoding.UTF8),
                Signature = signature
            };
            var response = GetLightningClient().VerifyMessage(request);
            return response;
        }

        public ChannelPoint OpenChannel(OpenChannelRequest request)
        {
            var response = GetLightningClient().OpenChannelSync(request);
           // var stream = response.ResponseStream;
          //  stream.MoveNext(CancellationToken.None).GetAwaiter().GetResult();
           // return stream.Current.ChanPending;
            return response;
        }

        public PendingChannelsResponse ListPendingChannels()
        {
            var request = new PendingChannelsRequest();
            var response = GetLightningClient().PendingChannels(request);
            return response;
        }

        public ListChannelsResponse ListChannels()
        {
            var request = new ListChannelsRequest();
            var response = GetLightningClient().ListChannels(request);
            return response;
        }

        public ChannelBalanceResponse ChannelBalance()
        {
            var request = new ChannelBalanceRequest();
            var response = GetLightningClient().ChannelBalance(request);
            return response;
        }

        public void CloseChannel(string channelPoint, bool force)
        {
            var request = new CloseChannelRequest
            {
                ChannelPoint = new ChannelPoint
                {
                    FundingTxidStr = channelPoint.Split(':')[0],
                    OutputIndex = uint.Parse(channelPoint.Split(':')[1])
                },
                Force = force
            };
            var stream = GetLightningClient().CloseChannel(request).ResponseStream;
            stream.MoveNext(CancellationToken.None);
        }

        public ClosedChannelsResponse ListClosedChannels()
        {
            var request = new ClosedChannelsRequest
            {
                Abandoned = true,
                Breach = true,
                Cooperative = true,
                FundingCanceled = true,
                LocalForce = true,
                RemoteForce = true
            };
            var response = GetLightningClient().ClosedChannels(request);
            return response;
        }

        public ListPaymentsResponse ListPayments()
        {
            var request = new ListPaymentsRequest();
            var response = GetLightningClient().ListPayments(request);
            return response;
        }

        public PayReq DecodePaymentRequest(string paymentRequest)
        {
            var request = new PayReqString {PayReq = paymentRequest};
            var response = GetLightningClient().DecodePayReq(request);
            return response;
        }

        public QueryRoutesResponse QueryRoutes(string pubkey, long amount, long maxFixedFee = 0, long maxPercentFee = 0, int finalCltvDelta = 144, int maxRoutes = 10)
        {
            var request = new QueryRoutesRequest
            {
                Amt = amount,
                FeeLimit =
                    maxFixedFee > 0 ? new FeeLimit { Fixed = maxFixedFee } : maxPercentFee > 0 ? new FeeLimit { Percent = maxPercentFee } : null,
                FinalCltvDelta = finalCltvDelta,
                NumRoutes = maxRoutes,
                PubKey = pubkey
            };
            var response = GetLightningClient().QueryRoutes(request);
            return response;
        }

        public IAsyncStreamReader<SendResponse> SendPayment(string paymentRequest, int timeout)
        {
            var deadline = DateTime.UtcNow.AddSeconds(timeout);
            var duplexPaymentStreaming = GetLightningClient().SendPayment(Metadata.Empty, deadline, CancellationToken.None);
            var request = new SendRequest {PaymentRequest = paymentRequest};
            duplexPaymentStreaming.RequestStream.WriteAsync(request);
            return duplexPaymentStreaming.ResponseStream;
        }

        public IAsyncStreamReader<SendResponse> SendToRoute(PayReq paymentRequest, List<Route> routes, int timeout)
        {
            var deadline = DateTime.UtcNow.AddSeconds(timeout);
            var duplexPaymentStreaming = GetLightningClient().SendToRoute(Metadata.Empty, deadline, CancellationToken.None);
            var request = new SendToRouteRequest
            {
                PaymentHash = ByteString.CopyFrom(paymentRequest.PaymentHash, Encoding.UTF8)
            };
            if (routes != null) request.Routes.Add(routes);
            duplexPaymentStreaming.RequestStream.WriteAsync(request);

            return duplexPaymentStreaming.ResponseStream;
        }

        public SendResponse SyncSendPayment(string paymentRequest)
        {
            var request = new SendRequest { PaymentRequest = paymentRequest };
            var deadline = DateTime.UtcNow.AddSeconds(30);
            var response = GetLightningClient().SendPaymentSync(request, deadline: deadline);
            return response;
        }

        public DeleteAllPaymentsResponse DeleteAllPayments()
        {
            var request = new DeleteAllPaymentsRequest();
            var response = GetLightningClient().DeleteAllPayments(request);
            return response;
        }

        public ListInvoiceResponse ListInvoices()
        {
            var request = new ListInvoiceRequest();
            var response = GetLightningClient().ListInvoices(request);
            return response;
        }

        public AddInvoiceResponse AddInvoice(string memo, long value)
        {
            var request = new Invoice
            {
                Memo = memo,
                Value = value,
                FallbackAddr = NewAddress(NewAddressRequest.Types.AddressType.NestedPubkeyHash).Address,
                Private = true
            };
            var response = GetLightningClient().AddInvoice(request);
            return response;
        }

        public IAsyncStreamReader<Invoice> SubscribeInvoices()
        {
            var request = new InvoiceSubscription();
            var response = GetLightningClient().SubscribeInvoices(request);
            return response.ResponseStream;
        }

        public ChannelGraph DescribeGraph()
        {
            var request = new ChannelGraphRequest
            {
                
            };
            var response = GetLightningClient().DescribeGraph(request);
            return response;
        }

        public ChannelEdge GetChannelEdge(ulong channelId)
        {
            var request = new ChanInfoRequest
            {
                ChanId = channelId
            };
            var response = GetLightningClient().GetChanInfo(request);
            return response;
        }

        public NodeInfo GetNodeInfo(string pub_key)
        {
            var request = new NodeInfoRequest
            {
                PubKey = pub_key
            };
            var response = GetLightningClient().GetNodeInfo(request);
            return response;
        }

        public NetworkInfo GetNetworkInfo()
        {
            var request = new NetworkInfoRequest();
            var response = GetLightningClient().GetNetworkInfo(request);
            return response;
        }
    }
}
