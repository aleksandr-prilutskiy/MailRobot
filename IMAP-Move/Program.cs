using System;
using GemBox.Email;
using GemBox.Email.Imap;

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
                Console.WriteLine("use: imap <command> <msg id> <folder from> <folder to>");
                Console.WriteLine("commands:");
                Console.WriteLine(" move <msg id> <folder from> <folder to>");
                return;
            }
            switch (args[0].ToLower())
            {
                case "move":
                    Move(args);
                    break;
            }

            //Console.ReadLine();
        }
        
        static void Move(string[] args)
        {
            if (args.Length < 4) return;
            if (!int.TryParse(args[1], out int id)) return;
            string folderFrom = args[2];
            string folderTo = args[3];
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
                imap.SelectFolder(folderFrom, false);
                imap.MoveMessage(id, folderTo);
                imap.Disconnect();
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
