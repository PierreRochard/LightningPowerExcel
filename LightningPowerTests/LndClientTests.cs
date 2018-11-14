using Grpc.Core;
using LightningPower;
using Lnrpc;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LightningPowerTests
{
    [TestClass()]
    public class LndClientIntegrationTests
    {
        [TestMethod()]
        public void UnlockWalletTestWrongPassword()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act and Assert
            Assert.ThrowsException<RpcException>(() => lndClient.UnlockWallet("wrong_password"));
        }

        [TestMethod()]
        public void UnlockWalletTestRightPassword()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act and Assert
            try
            {
                UnlockWalletResponse response = lndClient.UnlockWallet(new LndClientConfiguration().WalletPassword);
            }
            catch (RpcException e)
            {
                // Wallet is already unlocked
                Assert.AreEqual("unknown service lnrpc.WalletUnlocker", e.Status.Detail);
            }
        }

        [TestMethod()]
        public void GetInfoTest()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act
            GetInfoResponse response = lndClient.GetInfo();

            Assert.AreEqual("0.5.0-beta commit=3b2c807288b1b7f40d609533c1e96a510ac5fa6d", response.Version);
        }

        [TestMethod()]
        public void NewAddressTest()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act
            NewAddressResponse response = lndClient.NewAddress();

            Assert.IsNotNull(response.Address);
        }

        [TestMethod()]
        public void ListChannelsTest()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act
            ListChannelsResponse response = lndClient.ListChannels();

            Assert.IsNotNull(response);
        }

        [TestMethod()]
        public void ListPaymentsTest()
        {
            LndClient lndClient = new LndClient();
            var response = lndClient.ListPayments();
            Assert.IsNotNull(response);
        }

        [TestMethod()]
        public void SendPaymentTest()
        {
            LndClient lndClient = new LndClient();
            // Todo: query a testnet lapp for a payment request
            var response = lndClient.SendPayment("", 30);
            Assert.IsNotNull(response);
        }
    }
}