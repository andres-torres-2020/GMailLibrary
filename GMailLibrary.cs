using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ApplicationLogger;

namespace GMailLibrary
{
    public interface IGMailAttachmentProcessor
    {
        Boolean Process(string AttachmentFilename, MessagePartBody messagePartBody);
    }
    public class GMailLibrary
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json

        static readonly string UserMe = "me";
        readonly ILogger AppLogger;
        public GMailLibrary(ILogger appLogger)
        {
            AppLogger = appLogger;
        }
        public async Task<UserCredential> AuthorizeAsync(string CredentialsFileName, string CredentialsTokenPath, string[] ServiceScopes)
        {
            UserCredential credential;
            using (var stream =
                new FileStream(CredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = CredentialsTokenPath;
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    ServiceScopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true));
                AppLogger.Log("AuthorizeAsync: Credential file saved to: " + credPath);
            }
            return credential;
        }
        public GmailService GetService(UserCredential Credential, string ApplicationName)
        {
            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Credential,
                ApplicationName = ApplicationName
            });
            return service;
        }
        public async Task<List<Message>> GetMessages(GmailService Service, string searchString)
        {
            // Define parameters of request.
            UsersResource.MessagesResource.ListRequest request = Service.Users.Messages.List(UserMe);
            request.Q = searchString;

            // get Messages
            ListMessagesResponse response = await request.ExecuteAsync();
            List<Message> messages = new List<Message>();
            AppLogger.Log("\tMessages:");
            if (response.Messages == null || response.Messages.Count <= 0)
            {
                AppLogger.Log("\tNo messages found.");
            }
            while (response.Messages != null)
            {
                AppLogger.Log($"\tfound {response.Messages.Count} messages");
                messages.AddRange(response.Messages);

                if (!string.IsNullOrEmpty(response.NextPageToken))
                {
                    request = Service.Users.Messages.List(UserMe);
                    request.Q = searchString;
                    request.PageToken = response.NextPageToken;
                    response = await request.ExecuteAsync();
                }
                else
                {
                    break;
                }
            }
            AppLogger.Log($"\tgot {messages.Count} # messages");
            return messages;
        }
        public async Task<Message> GetFullMessage(GmailService Service, string MessageId)
        {
            AppLogger.Log($"email: {MessageId}");
            try
            {
                var fullMessageRequest = Service.Users.Messages.Get(UserMe, MessageId);
                var fullMessage = await fullMessageRequest.ExecuteAsync();
                return fullMessage;

            } catch (Exception ex)
            {
                AppLogger.Log(ex.Message);
            }
            return null;
        }
        public async Task<MessagePartBody> GetAttachment(GmailService service, Message message, string attachmentId)
        {
            return await service.Users.Messages.Attachments.Get(UserMe, message.Id, attachmentId).ExecuteAsync();
        }
        public async Task<Boolean> ProcessAttachments(GmailService service, List<Message> messages, IGMailAttachmentProcessor attachmentProcessor)
        {
            bool result = true;
            int count = 0;
            foreach (var email in messages)
            {
                AppLogger.Log($"[{++count}] GMailLibrary.ProcessAttachments");
                Message fullMessage = await GetFullMessage(service, email.Id);
                if (! await ProcessAttachments(service, fullMessage, attachmentProcessor))
                {
                    result = false;
                }
            }
            return result;
        }
        private async Task<Boolean> ProcessAttachments(GmailService service, Message email, IGMailAttachmentProcessor attachmentProcessor)
        {
            int count = 0;
            foreach (MessagePart part in email.Payload.Parts)
            {
                AppLogger.Log($"\tattachment part # {++count}");
                if (part.Filename.Length > 0)
                {
                    AppLogger.Log($"\tattachment: fileName=[{part.Filename}] partId={part.PartId}");
                    MessagePartBody messagePartBody = await GetAttachment(service, email, part.Body.AttachmentId);
                    //UsersResource.MessagesResource.AttachmentsResource.GetRequest attachmentRequest = service.Users.Messages.Attachments.Get(UserMe, email.Id, part.Body.AttachmentId);
                    //MessagePartBody messagePartBody = await attachmentRequest.ExecuteAsync();
                        //AppLogger.Log($"partSize={messagePartBody.Size}");
                    attachmentProcessor.Process(part.Filename, messagePartBody);
                    //AppLogger.Log();
                }
            }
            return true;
        }
    }
}
