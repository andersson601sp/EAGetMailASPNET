using System;

namespace EAGetMailASPNET
{
    class Program
    {
        static void Main(string[] args)
        {
            var mails = OutlookEmails.ReadEmailItems();
            int i = 1;

            foreach (var mail in mails)
            {
                Console.WriteLine("Mail no" + i);
                Console.WriteLine("Mail Receive from " + mail.EmailFrom);
                Console.WriteLine("Mail Subject " + mail.EmailSubject);
                Console.WriteLine("MailBody " + mail.EmailBody);
                Console.WriteLine("");
                i = i + 1;
            }
            Console.ReadKey();

        }
    }
}
