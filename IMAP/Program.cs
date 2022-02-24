using System;
using MailKit;
using MailKit.Net;

namespace IMAP
{
    class Program
    {
        private const string _iniFileName = "mail.ini";
        static string ServerAddr = "";
        static string ServerUser = "";
        static string ServerPassword = "";
        public static Common.IniFile IniFile;

        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("use: imap <command>");
                Console.WriteLine("commands:");
                Console.WriteLine(" folders");
                Console.WriteLine(" messages <folder>");
                Console.WriteLine(" move <ForReplyId> <folder from> <folder to>");
                return;
            }
            switch (args[0].ToLower())
            {
                case "folders":
                    Folders(args);
                    break;
                case "messages":
                    Messages(args);
                    break;
                case "move":
                    Move(args);
                    break;
            }
            //Console.ReadLine();
        }
        
        static void Folders(string[] args)
        {
            LoadIniFile();
            Console.WriteLine(ServerUser);
            try
            {
                Console.WriteLine("Connect to: " + ServerAddr);
                var client = new MailKit.Net.Imap.ImapClient();
                client.Connect(ServerAddr);
                Console.WriteLine("Connected.");
                client.Authenticate(ServerUser, ServerPassword);
                Console.WriteLine("Authenticated.");
                Console.WriteLine("Mail folders:");
                foreach (var folder in client.GetFolders(client.PersonalNamespaces[0]))
                {
                    Console.WriteLine(">" + folder);
                }
                client.Disconnect(true);
                Console.WriteLine("OK");
            }
            catch (Exception exeption)
            {
                Console.WriteLine("Error: " + exeption.Message);
            }
        }

        static void Messages(string[] args)
        {
            if (args.Length < 2) return;
            string folderName = args[1];
            LoadIniFile();
            Console.WriteLine(ServerUser);
            try
            {
                Console.WriteLine("Connect to: " + ServerAddr);
                var client = new MailKit.Net.Imap.ImapClient();
                client.Connect(ServerAddr);
                Console.WriteLine("Connected.");
                client.Authenticate(ServerUser, ServerPassword);
                Console.WriteLine("Authenticated.");
                Console.WriteLine("Messages in folder: '" + folderName + "'");
                IMailFolder folder = client.GetFolder(folderName);
                folder.Open(FolderAccess.ReadWrite);
                for (int i = 0; i < folder.Count; i++)
                {
                    MimeKit.MimeMessage message = folder.GetMessage(i);
                    Console.WriteLine("#" + i.ToString() + ": " + message.From);
                }
                client.Disconnect(true);
                Console.WriteLine("OK");
            }
            catch (Exception exeption)
            {
                Console.WriteLine("Error: " + exeption.Message);
            }
        }

        static void Move(string[] args)
        {
            if (args.Length < 4) return;
            string id = args[1];
            string folderFrom = args[2];
            string folderTo = args[3];
            LoadIniFile();
            Console.WriteLine(ServerUser);
            try
            {
                Console.WriteLine("Connect to: " + ServerAddr);
                var client = new MailKit.Net.Imap.ImapClient();
                client.Connect(ServerAddr);
                Console.WriteLine("Connected.");
                client.Authenticate(ServerUser, ServerPassword);
                Console.WriteLine("Authenticated.");
                Console.WriteLine("Messages in folder: '" + folderFrom + "'");
                IMailFolder folder = client.GetFolder(folderFrom);
                IMailFolder folder2 = client.GetFolder(folderTo);
                folder.Open(FolderAccess.ReadWrite);
                for (int i = 0; i < folder.Count; i++)
                {
                    MimeKit.MimeMessage message = folder.GetMessage(i);
                    if (message.MessageId == id)
                    {
                        folder.MoveTo(i, folder2);
                        Console.WriteLine("Move to: '" + folderTo + "'");
                    }
                }
                client.Disconnect(true);
                Console.WriteLine("OK");
            }
            catch (Exception exeption)
            {
                Console.WriteLine("Error: " + exeption.Message);
            }
        }

        static void LoadIniFile()
        {
            IniFile = new Common.IniFile(_iniFileName);
            ServerAddr = IniFile.ReadString("Mail", "Server", "");
            ServerUser = IniFile.ReadString("Mail", "User", "");
            ServerPassword = IniFile.ReadString("Mail", "Password", "");
        } // LoadIniFile()
    }
}
