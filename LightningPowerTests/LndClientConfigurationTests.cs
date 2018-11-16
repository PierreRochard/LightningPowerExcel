using Microsoft.VisualStudio.TestTools.UnitTesting;
using LightningPower;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightningPowerTests
{
    [TestClass]
    public class LndClientConfigurationTests
    {
        public LndClientConfiguration Config;

        [TestInitialize]
        public void Setup()
        {
            Config = new LndClientConfiguration();
        }

        [TestMethod]
        public void CaCertPathTest()
        {
            var path = Config.CaCertPath;
            Assert.IsTrue(path.EndsWith("tls.cert"));
        }

        [TestMethod]
        public void SslCredentialsTest()
        {
            var sslCredentials = Config.SslCredentials;
            Assert.IsTrue(sslCredentials.RootCertificates.StartsWith("-----BEGIN CERTIFICATE-----"));

        }

        [TestMethod]
        public void MacaroonStringTest()
        {
            var macaroon = Config.MacaroonString;
            Assert.IsNull(macaroon);
        }
    }
}
