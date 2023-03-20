using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

namespace GraphTutorial
{
    class GraphHelper
    {
        private static Settings _settings;
        // User auth token credential
        private static DeviceCodeCredential _deviceCodeCredential;
        // Client configured with user authentication
        private static GraphServiceClient _userClient;

        public static void InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {
            _settings = settings;

            _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
                settings.AuthTenant, settings.ClientId);

            _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
        }

        public static async Task<string> GetUserTokenAsync()
        {
            // Ensure credential isn't null
            _ = _deviceCodeCredential ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            // Ensure scopes isn't null
            _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

            // Request token with given scopes
            var context = new TokenRequestContext(_settings.GraphUserScopes);
            var response = await _deviceCodeCredential.GetTokenAsync(context);
            return response.Token;
        }


        // App-ony auth token credential
        private static ClientSecretCredential _clientSecretCredential;
        // Client configured with app-only authentication
        private static GraphServiceClient _appClient;

        public static void EnsureGraphForAppOnlyAuth(Settings settings)
        {
            // Ensure settings isn't null
            _settings = settings;

            _ = _settings ??
                throw new System.NullReferenceException("Settings cannot be null");

            if (_clientSecretCredential == null)
            {
                _clientSecretCredential = new ClientSecretCredential(
                    _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
            }

            if (_appClient == null)
            {
                _appClient = new GraphServiceClient(_clientSecretCredential,
                    // Use the default scope, which will request the scopes
                    // configured on the app registration
                    new[] { "https://graph.microsoft.com/.default" });
            }
        }

        public static IMailFolderMessagesCollectionPage GetInbox()
        {
            // Ensure client isn't null
            _ = _appClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return _appClient.Me
                // Only messages from Inbox folder
                .MailFolders["Inbox"]
                .Messages
                .Request()
                .Select(m => new
                {
                    // Only request specific properties
                    m.From,
                    m.IsRead,
                    m.ReceivedDateTime,
                    m.Subject
                })
                // Get at most 25 results
                .Top(25)
                // Sort by received time, newest first
                .OrderBy("ReceivedDateTime DESC")
                .GetAsync()
                .Result;
        }

        public static async Task SendMailAsync(string subject, string body, string recipient)
        {
            // Ensure client isn't null
            _ = _appClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            // Create a new message
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    Content = body,
                    ContentType = BodyType.Text
                },
                ToRecipients = new Recipient[]
                {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }
                }
            };

            //  Send the message
            await _appClient.Me
                .SendMail(message)
                .Request()
                .PostAsync();

            //var meRequest = _appClient.Me;
            //var sendMailRequest = meRequest.SendMail(message);
            //var request = sendMailRequest.Request();
            // request.PostAsync();
        }


        //public static void SendMail(string subject, string body, string recipient)
        //{
        //    // Ensure client isn't null
        //    _ = _appClient ??
        //        throw new System.NullReferenceException("Graph has not been initialized for user auth");

        //    // Create a new message
        //    var message = new Message
        //    {
        //        Subject = subject,
        //        Body = new ItemBody
        //        {
        //            Content = body,
        //            ContentType = BodyType.Text
        //        },
        //        ToRecipients = new Recipient[]
        //        {
        //    new Recipient
        //    {
        //        EmailAddress = new EmailAddress
        //        {
        //            Address = recipient
        //        }
        //    }
        //        }
        //    };
        //    var meRequest = _appClient.Me;
        //    var sendMailRequest =  meRequest.SendMail(message);
        //    var request = sendMailRequest.Request();
        //    var post = request.PostAsync(); 

        //    // Send the message

        //}
    }
}
