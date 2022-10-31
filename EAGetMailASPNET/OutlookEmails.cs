using EAGetMail;
using System;
using System.Collections.Generic;
using System.Text;

namespace EAGetMailASPNET
{
    public class OutlookEmails
    {
        public string EmailFrom { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody { get; set; }

        /// <summary>
        /// Outlook Office Port 143 - Server Name: outlook.office365.com
        /// Gmail Port 993 - Server Name: imap.gmail.com
        /// Outlook Port 993 - Server Name: outlook.live.com
        /// </summary>
        /// <returns></returns>
        public static List<OutlookEmails> ReadEmailItems()
        {
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            var server = "outlook.live.com";
            var user = "yourmail";
            var pwrd = "yourpassword";
            MailServer oServer = new MailServer(server, user, pwrd, ServerProtocol.Imap4);
            MailClient oClient = new MailClient("TryIt");
            oServer.SSLConnection = true;
            oServer.Port = 993;

            try
            {
                oClient.Connect(oServer);
                MailInfo[] infos = oClient.GetMailInfos();
                Console.WriteLine(infos.Length);
                for (int i = 0; i < infos.Length; i++)
                {
                    MailInfo info = infos[i];
                    Mail oMail = oClient.GetMail(info);


                    listEmailDetails.Add( new OutlookEmails { EmailBody = oMail.TextBody, EmailFrom = oMail.From.ToString() , EmailSubject = oMail.Subject } );

                    //Console.WriteLine("From: {0}", oMail.From.ToString());
                    //oClient.Delete(info);
                }
                oClient.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return listEmailDetails;
        }
    }
}
