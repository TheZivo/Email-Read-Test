using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Read_Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var mails = OutlookEmails.ReadMailItems();
            //TEMP Iterator for test
            int i = 1;
            foreach (var mail in mails)
            { 
                //TEMP :: printing to console as test 
                Console.WriteLine("Mail No " + i);
                Console.WriteLine("Mail Recieved From: " + mail.EmailFrom);
                Console.WriteLine("Mail Subject: " + mail.EmailSubject);
                Console.WriteLine("");
                i = i + 1;
                //TEMP

            }
            Console.ReadKey();
        }
    }
}

