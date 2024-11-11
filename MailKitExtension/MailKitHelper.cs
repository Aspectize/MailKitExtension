using MailKit;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MailKitExtention {

    public class MailKitHelper {

        static Regex rxCapitalize = new Regex(@"\b(\w)");
        static string Capitalize(string input) {

            return rxCapitalize.Replace(input, m => m.Groups[1].Value.ToUpper());
        }
        static string evalParams(string content, Dictionary<string, string> contentParams) {

            if (contentParams != null) {
                foreach (var key in contentParams.Keys) {

                    var rxKey = new Regex($"%{key}%");
                    var kValue = contentParams[key];

                    content = rxKey.Replace(content, kValue);
                }
            }

            return content;
        }

        public static void ThrowExceptionIfInvalidAddresses(params string[] emailAddresses) {

            if (emailAddresses != null) {

                var ivAddresses = new List<string>();

                foreach (var email in emailAddresses) {

                    if (!IsValid(email)) ivAddresses.Add(email);
                }

                if (ivAddresses.Count > 0) {

                    var all = string.Join(", ", ivAddresses);
                    throw new Aspectize.Core.SmartException($"Invalid email adresses : {all}");
                }
            }

        }

        public static bool IsValid(string email) {
            MailboxAddress x;

            return MailboxAddress.TryParse(email, out x);
        }

        public static void Send(string smtpServer, int port, EnumMailKiAuthenticationType authenticationType, string login, string openIdTokenOrPassword, string senderEMail, string senderDisplay, string[] receivers, string title, string htmlBody, Dictionary<string, byte[]> attachments = null, string[] bcc = null, string[] cc = null, Dictionary<string, string> contentParams = null, bool withReceipt = false) {

            ThrowExceptionIfInvalidAddresses(senderEMail);
            ThrowExceptionIfInvalidAddresses(receivers);
            ThrowExceptionIfInvalidAddresses(cc);
            ThrowExceptionIfInvalidAddresses(bcc);

            title = evalParams(title, contentParams);
            var body = evalParams(htmlBody, contentParams);

            var bodyBuilder = new MimeKit.BodyBuilder();

            bodyBuilder.HtmlBody = body;

            if (attachments != null) {
                foreach (var aName in attachments.Keys) {
                    bodyBuilder.Attachments.Add(aName, attachments[aName]);
                }
            }

            var mailMessage = new MimeKit.MimeMessage();

            mailMessage.Subject = title;
            mailMessage.Body = bodyBuilder.ToMessageBody();

            mailMessage.Sender = new MailboxAddress(Encoding.UTF8, senderDisplay, senderEMail);
            mailMessage.From.Add(new MailboxAddress(Encoding.UTF8, senderDisplay, senderEMail));

            foreach (var receiver in receivers) {

                var display = Capitalize(receiver.Split('@')[0].Replace('.', ' '));
                mailMessage.To.Add(new MailboxAddress(Encoding.UTF8, display, receiver.Trim()));
            }
            if (cc != null) {
                foreach (var receiver in cc) {
                    var display = Capitalize(receiver.Split('@')[0].Replace('.', ' '));
                    mailMessage.Cc.Add(new MailboxAddress(Encoding.UTF8, display, receiver.Trim()));
                }
            }
            if (bcc != null) {
                foreach (var receiver in bcc) {
                    var display = Capitalize(receiver.Split('@')[0].Replace('.', ' '));
                    mailMessage.Bcc.Add(new MailboxAddress(Encoding.UTF8, display, receiver.Trim()));
                }
            }

            if (withReceipt) {
                mailMessage.Headers.Add("Disposition-Notification-To", senderEMail);
                mailMessage.Headers.Add("Return-Receipt-To", senderEMail);
            }


            using (var client = new MailKit.Net.Smtp.SmtpClient()) {

                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect(smtpServer, port, SecureSocketOptions.StartTlsWhenAvailable);

                //client.AuthenticationMechanisms.Clear();
                //client.AuthenticationMechanisms.Add("XOAUTH2");
                switch (authenticationType) {
                    case EnumMailKiAuthenticationType.SimplePasswword:
                        client.Authenticate(login, openIdTokenOrPassword);

                        break;
                    case EnumMailKiAuthenticationType.OpenIdAccessToken:
                        client.Authenticate(new SaslMechanismOAuth2(login, openIdTokenOrPassword));

                        break;
                    case EnumMailKiAuthenticationType.ClientCertificat:
                        break;
                }
                client.Send(mailMessage);
                client.Disconnect(true);
            }
        }

        //public static void ImapRead(string imapServer, string login, string password, int port = 993) {

        //    using (var client = new MailKit.Net.Imap.ImapClient ()) {

        //        client.Connect(imapServer, port, true);
        //        client.Authenticate(login, password, );

        //        var inbox = client.Inbox;
        //        inbox.Open(FolderAccess.ReadOnly);

        //        var doneFolder = client.GetFolder("Done") ?? client.Inbox.Create("Done", true);

        //        var recent = inbox.Recent;

        //        var recentMessages = inbox.Search(SearchQuery.DeliveredAfter(DateTime.Today));
        //        var notDone = inbox.Search(SearchQuery.NotSeen);
        //        //for (int i = Math.Max(0, inbox.Count - 10); i < inbox.Count; i++) {
        //        foreach (var uid in notDone) {
        //            var message = inbox.GetMessage(uid);
        //            var from = message.From;
        //            var subject = message.Subject;
        //            var date = message.Date;
        //            var body = message.TextBody;

        //            inbox.AddFlags(uid, MessageFlags.Seen, true);
        //            inbox.MoveTo(uid, doneFolder);
        //        }
        //    }
        //}
    }
}
