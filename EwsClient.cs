using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;

namespace EwsOAuth
{
    class EwsService
    {
        public static ExchangeService EwsClient(string identity, string authToken)
        {
            string ewsUri = " https://outlook.office365.com/EWS/Exchange.asmx";

            ExchangeService ewsClient = new ExchangeService
            {
                Url = new Uri(ewsUri),
                Credentials = new OAuthCredentials(authToken),
                ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, identity),
            };
            ewsClient.HttpHeaders.Add("X-AnchorMailbox", identity);

            return ewsClient;
        }
    }
}
