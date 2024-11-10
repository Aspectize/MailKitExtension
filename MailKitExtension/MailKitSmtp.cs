using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Aspectize.Core;
using System.Security.Permissions;

namespace MailKitExtention {

    public enum EnumMailKiAuthenticationType {
        SimplePasswword,
        OpenIdAccessToken,
        ClientCertificat
    }

    [Service(Name = "MailKitSmtp", ConfigurationRequired = true)]
    public class MailKitSmtp : IAspectizeSMTPService, ISingleton //, IInitializable, ISingleton
    {

        [Parameter]
        string Host;

        [Parameter]
        int Port = 587; // 25 or 465 (SSL) or 587 (StartTLS)

        [Parameter]
        bool Ssl = false;

        [Parameter]
        string Login;

        [Parameter]
        string Password;

        [Parameter(DefaultValue = "info@mydomain.com")]
        string Expediteur;

        [Parameter]
        string ExpediteurDisplay;

        string[] getDestinataires(bool sendCopyToExpediteur, string destinataire) {

            var destinataires = new List<string>();
            destinataires.Add(destinataire);
            if (sendCopyToExpediteur) destinataires.Add(Expediteur);


            return destinataires.ToArray();
        }
        string[] getDestinataires(bool sendCopyToExpediteur, string[] destinataires) {

            if (sendCopyToExpediteur) {

                var all = new List<string>(destinataires);
                all.Add(Expediteur);

                return all.ToArray();

            } else return destinataires;
        }

        void IAspectizeSMTPService.SendMail(bool sendCopyToExpediteur, string[] destinataires, string subject, string message, Dictionary<string, byte[]> dicoAttachements) {

            destinataires = getDestinataires(sendCopyToExpediteur, destinataires);

            MailKitHelper.Send(Host, Port, EnumMailKiAuthenticationType.SimplePasswword, Login, Password, Expediteur, ExpediteurDisplay, destinataires, subject, message, dicoAttachements);
        }

        void IAspectizeSMTPService.SendMailFrom(string expediteur, string[] destinataires, string subject, string message) {

            if (!MailKitHelper.IsValid(expediteur)) {

                expediteur = Expediteur;
            } 

            MailKitHelper.Send(Host, Port, EnumMailKiAuthenticationType.SimplePasswword, Login, Password, expediteur, ExpediteurDisplay, destinataires, subject, message);
        }

        void IAspectizeSMTPService.SendMailSimple(bool sendCopyToExpediteur, string destinataire, string subject, string message) {

            var destinataires = getDestinataires(sendCopyToExpediteur, destinataire);

            MailKitHelper.Send(Host, Port, EnumMailKiAuthenticationType.SimplePasswword, Login, Password, Expediteur, ExpediteurDisplay, destinataires, subject, message);
        }

        void IAspectizeSMTPService.SendMailWithBcc(bool sendCopyToExpediteur, string[] destinataires, string subject, string message, Dictionary<string, byte[]> dicoAttachements, string[] bcc) {

            destinataires = getDestinataires(sendCopyToExpediteur, destinataires);

            MailKitHelper.Send(Host, Port, EnumMailKiAuthenticationType.SimplePasswword, Login, Password, Expediteur, ExpediteurDisplay, destinataires, subject, message, dicoAttachements, bcc);
        }
    }

}
