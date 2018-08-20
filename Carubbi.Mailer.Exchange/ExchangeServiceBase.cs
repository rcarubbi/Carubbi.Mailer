using Carubbi.ServiceLocator;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Net;

namespace Carubbi.Mailer.Exchange
{
    public abstract class ExchangeServiceBase
    {
        protected AppSettings Config;

        protected ExchangeServiceBase()
        {
            Config = new AppSettings("CarubbiMailer");
        }

        protected ExchangeService Instance = null;
        
        protected ExchangeService GetExchangeService(string username, string password)
        {
                if (Instance == null)
                {
                    SetupExchangeClient(username, password);
                }

                return Instance;
        }

        protected void SetupExchangeClient(string username, string password)
        {
            ServicePointManager.ServerCertificateValidationCallback =
                    delegate
                    {
                        return true;
                    };

            InitializeClient(username, password);
        }

        protected void InitializeClient(string username, string password)
        {
            var ewsUrl = Config["URL_SERVICO_EXCHANGE"];

            Instance = new ExchangeService((ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), Config["VERSAO_EXCHANGE"]), TimeZoneInfo.Local);
            if (!string.IsNullOrEmpty(Config["WEB_PROXY"]))
            {
                var wp = new WebProxy(Config["WEB_PROXY"]);
                Instance.WebProxy = wp;

            }
            Instance.Credentials = new NetworkCredential(username, password);
            Instance.Url = new Uri(ewsUrl);
        }
    }

}