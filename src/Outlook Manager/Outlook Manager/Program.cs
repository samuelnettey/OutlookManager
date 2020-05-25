using System;

namespace Outlook_Manager
{
    class Program
    {
        static void Main(string[] args)
        {
            var mails = OutlookMail.ReadMails();

            int i = 0;
            foreach (var mail in mails)
            {
                Console.WriteLine("Mail No "+  i);
                Console.WriteLine("Received From "+ mail.From );
                Console.WriteLine("Mail Subject "+  mail.Subject);
                Console.WriteLine("Mail Body "+  mail.Body);
                Console.WriteLine("");
                i++;
            }
        }
    }
}