using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Email;
using GemBox.Email.Imap;

namespace IMAP
{
    class Program
    {
        private const string _iniFileName = "MailRobot.ini";    // Имя файла настроек сервиса
        static string SaveDir = System.IO.Directory.GetCurrentDirectory() + "\\mail";
        static string ServerAddr = "";
        static string ServerUser = "";
        static string ServerPassword = "";
        public static Common.IniFile IniFile;                   // Объект для работы с файлом конфигурации программы

        static void Main(string[] args)
        {
            LoadIniFile();
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            Console.WriteLine(ServerUser);
            try
            {
                ImapClient imap = new ImapClient(ServerAddr);
                imap.ConnectTimeout = TimeSpan.FromSeconds(4);
                imap.Connect();
                Console.WriteLine("Connected.");
                imap.Authenticate(ServerUser, ServerPassword);
                Console.WriteLine("Authenticated.");

                Console.WriteLine("Folders:");
                imap.SelectFolder("Inbox", false);
                var folders = imap.ListFolders();
                foreach (var folder in folders)
                {
                    ImapFolderStatus status = imap.GetFolderStatus(folder.Name);
                    Console.WriteLine(folder.Name + " > " + status.IsReadOnly + " / " + status.PermanentFlags.Count);
                }

                imap.SelectFolder("INBOX.Robot", false);
                var messages = imap.ListMessages();
                foreach (var id in messages)
                {
                    MailMessage message = imap.GetMessage(id.Number);
                    Console.WriteLine(id.Number + ": From:" + message.From);
                }

                //if (messages.Count > 0)
                //{
                //    imap.MoveMessage(messages[0].Number, "INBOX.Junk");
                //}

                imap.Disconnect();
                Console.WriteLine("OK");
            }
            catch (Exception exeption)
            {
                Console.WriteLine("Error: " + exeption.Message);
            }
            Console.ReadLine();
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
