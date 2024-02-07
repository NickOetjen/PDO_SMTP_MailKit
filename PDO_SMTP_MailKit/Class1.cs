using System;

using MailKit.Net.Smtp;
using MailKit;
using MimeKit;


//? err.Description
// MimeKit 4.3 => Die Datei oder Assembly "System.Runtime.CompilerServices.Unsafe, Version=4.0.4.1, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" oder eine Abhängigkeit davon wurde nicht gefunden.Das System kann die angegebene Datei nicht finden.
// => MimeKit 3.4.2 mit geringeren Anforderungen zu System.Runtime.CompilerServices.Unsafe (=> downgrade auf Version 4.5.2) installiert



namespace PDO_SMTP_MailKit
{
    public class SendMail_using_SMTP
    {
        // PDO: Configured as COM-Addin. Can be used as AddIn in VBA , after installation and inclusion as reference.

        // Background: From VBA status quo 2023 mostly the cdo-library is used, which is 
        // deprecated / not supported by newer Exchange-Servers (>= 2019), and StartTLS is not supported (only with a proxy solution). 
        // Furthermore, the common SMTPClient for C# is not recommended for new development by 2024.

        // To register it on your own machine, run VS as administrator
        // To register a COMAddIn on another machine, see regasm.exe


        private string Version = "20240111";
        private string strResult = "";

        // E-Mail-Configuration

        private MimeMessage message = new MimeMessage();
        private Multipart multipart = new Multipart("mixed");   // Needed to assemble the amil from parts, e.g. body and attachments


        private string SetsmtpServer = "";
        private int SetForceSmtpPort = 0; // mit ssl:  465; // mit StartTLS: 587 // plaintext: 25;  Will normally be set using SetIsSSL 
        private bool SetIsSSL = true;      // PDO: This enables StartTLS . 
        private string SetPassword = "";
        private string SetUser = "";        // PDO: If left blank SetEmailFrom will be used

        private string SetEmailFrom = "";  // default, sollte bei Aufruf überschrieben werden

        private string SetSubject = "New Test-E-Mail";



        //TextPart for multipart
        TextPart Sethtmlbody = new TextPart("html")
        {
            Text = "Can I <strong>test</strong> you?"
        };




        // Methods to set values
        public void smtpServer(string str)
        {
            SetsmtpServer = str;
        }

        public void IsSSL(bool boo)
        {
            SetIsSSL = boo;
        }

        public void ForceSmtpPort(int i)
        {
            SetForceSmtpPort = i;
        }

        public void Password(string str)
        {
            SetPassword = str;
        }

        public void User(string str)
        {
            SetUser = str;
        }

        public void Subject(string str)
        {
            SetSubject = str;
        }

        public void EmailFrom(string str)
        {
            SetEmailFrom = str;
        }

        public void To_Add(string str)
        {
            // PDO: Multiple Recipients separated by  ;
            string[] str_Shredder = str.Split(';');
            string s_trim;

            foreach (string s in str_Shredder)
            {

                s_trim = s.Trim();

                try
                {
                    //mail.To.Add(str);
                    message.To.Add(new MailboxAddress(s_trim, s_trim));
                }
                catch (Exception ex)
                {
                    strResult = "PDO__SMTP_MailKit: The recipient-address is invalid. Please check. Error: " + ex.ToString();
                }

            }
        }

        public void CC_Add(string str)
        {
            // PDO: Multiple Recipients separated by  ;
            string[] str_Shredder = str.Split(';');
            string s_trim;

            foreach (string s in str_Shredder)
            {

                s_trim = s.Trim();

                try
                {
                    //mail.CC.Add(str);
                    message.Cc.Add(new MailboxAddress(s_trim, s_trim));
                }
                catch (Exception ex)
                {
                    strResult = "PDO_SMTP_MailKit: The CC-address is invalid. Please check. Error: " + ex.ToString();
                }

            }
        }

        public void BCC_Add(string str)
        {
            // PDO: Multiple Recipients separated by  ;
            string[] str_Shredder = str.Split(';');
            string s_trim;

            foreach (string s in str_Shredder)
            {

                s_trim = s.Trim();

                try
                {
                    //mail.Bcc.Add(str);
                    message.Bcc.Add(new MailboxAddress(s_trim, s_trim));
                }
                catch (Exception ex)
                {
                    strResult = "PDO_SMTP_MailKit: The Bcc-address is invalid. Please check. Error: " + ex.ToString();
                }
            
            }
        }

        public void HTMLBody(string str)
        {
            Sethtmlbody.Text = str;
            multipart.Add(Sethtmlbody);
        }

        public void Attachment_Add(string str)
        {

            //AttachmentPart for multipart
            //MimePart Attachment = new MimePart("image", "jpg")
            MimePart Attachment = new MimePart("image", str.Substring(str.LastIndexOf(".") + 1))
            {
                Content = new MimeContent(new System.IO.FileStream(str, System.IO.FileMode.Open)),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = str.Substring(str.LastIndexOf("/") + 1)   // "Beispiel_jpg_Datei.jpg"
            };

            multipart.Add(Attachment);

        }



        // Just to check for connection to COM 
        public string HelloWorld()
        {
            return "Hello World, this is PDO_SMTP_MailKit Version " + Version;
        }




        // Mailing function
        public string SendMail()
        {

            // Configure Mail with some defaults, not using the methods, to make it more tolerant to faults

            if (string.IsNullOrEmpty(SetEmailFrom))
            {
                throw new Exception("PDO_COMAddIn_SMTP_MailKit: The sender-address must not be empty. Please check.");

            }

            if (SetForceSmtpPort == 0) { SetForceSmtpPort = (SetIsSSL) ? 587 : 465; }



            //var message = new MimeMessage();

            message.From.Add(new MailboxAddress(SetEmailFrom, SetEmailFrom));      //message.From.Add(new MailboxAddress("PDO", "nick.oetjen@mlpdialog.de"));

            // Empfänger werden per Methode gesetzt
            //message.To.Add(new MailboxAddress("nick", "nick.oetjen@mlpdialog.de"));

            message.Subject = SetSubject;



            //var Attachment = new MimePart("image", "jpg")
            //{
            //    Content = new MimeContent(new System.IO.FileStream("M:/Telefonie/MLP/Telefonie_DB/Controlling/Nick/__Desktop/Beispiel_jpg_Datei.jpg", System.IO.FileMode.Open)),
            //    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
            //    ContentTransferEncoding = ContentEncoding.Base64,
            //    FileName = "Beispiel_jpg_Datei.jpg"
            //};

            //multipart.Add(Attachment);


            // PDO: Body is assembled from parts
            message.Body = multipart;

            //message.Body = new TextPart("plain")
            //{
            //    Text = @"Hey Me,

            //            I just wanted to let you know that Monica and I were going to go play some football, you in?

            //            -- Me"
            //};



            using (var client = new SmtpClient())
            {
                // Send E-Mail , when nothing went wrong so far
                if (strResult == "")
                {
                    try
                    {


                        //client.Connect("smtp.strato.de", 587, false);
                        //client.Authenticate("oetjen@pdoetjen.de", "x");



                        client.Connect(SetsmtpServer, SetForceSmtpPort, (SetIsSSL == true) ? MailKit.Security.SecureSocketOptions.StartTls : MailKit.Security.SecureSocketOptions.Auto);  // Hier als letztes den Encryption-Mode  //client.Connect("smtp.mlpdialog.de", 25, MailKit.Security.SecureSocketOptions.StartTls);  // 587 evt. besser, weil das der dafür vorgesehene Port ist

                        // Note: only needed if the SMTP server requires authentication
                        client.Authenticate((SetUser == "") ? SetEmailFrom : SetUser, SetPassword);    //client.Authenticate("mlp-software", "H39bih3D6a&3Q!X0%=o?pnzh#");   //client.Authenticate("mlp-web-01", "fHVx!?H2Y7N9jdu+3H");



                        client.Send(message);

                        strResult = "Ok";

                    }
                    catch (Exception ex)
                    {

                        strResult = ex.ToString();

                    }

                }

                client.Disconnect(true);
            }


            // CleanUp
            message.Dispose();
            Sethtmlbody.Dispose();
            multipart.Dispose();    // To free the files



            return strResult == "" ? "(No return message)" : strResult;


            // ****************************************
            // C|:-)  www.pdo.digital says Thank you !
            // ****************************************


        }


    }
}
