using System;
using System.Collections.Generic;

namespace MailRobot
{
    class Program
    {
        private const string _iniFileName = "MailRobot.ini";    // Имя файла настроек сервиса
        static string SaveDir = System.IO.Directory.GetCurrentDirectory() + "\\mail";

        static string ServerAddr = "";
        static string ServerUser = "";
        static string ServerPassword = "";
        static string AppStart = "";
        static string AppMessage = "";
        static string AppFinish = "";

        static string Filters = "";
        
        public static Common.IniFile IniFile;                   // Объект для работы с файлом конфигурации программы

        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            LoadIniFile();
            PrepareDir(SaveDir);
            OpenPop.Pop3.Pop3Client pop3 = new OpenPop.Pop3.Pop3Client();
            if (!ConnectToMail(pop3)) return;
            ExecuteApp(AppStart);
            var n = pop3.GetMessageCount();
            Console.WriteLine("=" + n);
            for (int i = 1; i <= n; i++)
            {
                OpenPop.Mime.Message message = pop3.GetMessage(i);
                if (Filters.Contains("today"))
                    if ((DateTime.Today - message.Headers.DateSent).Days > 0)
                        continue;
                CleanDir(SaveDir);
                Console.WriteLine("Message: #" + i);
                Console.WriteLine("From: " + message.Headers.From);
                string to = "";
                foreach (var a in message.Headers.To)
                    to += a.Address + ", ";
                Console.WriteLine("To: " + to.Substring(0, to.Length - 2));
                Console.WriteLine("Date: " + message.Headers.Date);
                Console.WriteLine("DateSent: " + message.Headers.DateSent);
                Console.WriteLine("Subject: " + message.Headers.Subject);
                SaveFile("message.eml", message.RawMessage);
                List<OpenPop.Mime.MessagePart> attachments = message.FindAllAttachments();
                if (attachments != null)
                {
                    Console.WriteLine("Attachments:");
                    foreach (OpenPop.Mime.MessagePart attachment in attachments)
                    {
                        Console.WriteLine(">" + attachment.FileName);
                        SaveFile(attachment.FileName, attachment.Body);
                    }
                }
                ExecuteApp(AppMessage);
                Console.WriteLine();
            }
            pop3.Disconnect();
            CleanDir(SaveDir);
            ExecuteApp(AppFinish);
            Console.WriteLine("ok");
        } // Main(string[])

        /// <summary>
        /// 
        /// </summary>
        static void LoadIniFile()
        {
            IniFile = new Common.IniFile(_iniFileName);
            ServerAddr = IniFile.ReadString("Mail", "Server", "");
            ServerUser = IniFile.ReadString("Mail", "User", "");
            ServerPassword = IniFile.ReadString("Mail", "Password", "");
            Filters = IniFile.ReadString("Mail", "Filters", "");
            AppStart = IniFile.ReadString("App", "Start", "");
            AppMessage = IniFile.ReadString("App", "Message", "");
            AppFinish = IniFile.ReadString("App", "Finish", "");
        } // LoadIniFile()

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pop3"></param>
        /// <returns></returns>
        static bool ConnectToMail(OpenPop.Pop3.Pop3Client pop3)
        {
            try
            {
                pop3.Connect(ServerAddr, 110, false);
                pop3.Authenticate(ServerUser, ServerPassword);
            }
            catch (Exception e)
            {
                Console.WriteLine("error: " + e.Message);
                return false;
            }
            return true;
        } // ConnectToMail(OpenPop.Pop3.Pop3Client)

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dir"></param>
        static void PrepareDir(string dir)
        {
            if (!System.IO.Directory.Exists(dir))
                System.IO.Directory.CreateDirectory(dir);
        } // PrepareDir(string)

        /// <summary>
        /// Очистка каталога временного размещения файлов сообщений
        /// </summary>
        /// <param name="dir"></param>
        static void CleanDir(string dir)
        {
            string[] filelist = System.IO.Directory.GetFiles(dir, "*.*");
            foreach (string filename in filelist)
                System.IO.File.Delete(filename);
        } // CleanDir(string)

        /// <summary>
        /// Создание файла и запись двоичных данных в этот файл
        /// </summary>
        /// <param name="filename">Имя файла</param>
        /// <param name="data">Массив байт для записи в файл</param>
        static void SaveFile(string filename, byte[] data)
        {
            filename = SaveDir + "\\" + filename;
            try
            {
                System.IO.FileStream fstream = new System.IO.FileStream(filename, System.IO.FileMode.Create);
                fstream.Write(data, 0, data.Length);
                fstream.Close();
            }
            catch
            {
                Console.WriteLine("Error write to file: " + filename);
                return;
            }
        } // SaveFile(string, byte[])

        /// <summary>
        /// Запуск исполняемого файла
        /// </summary>
        /// <param name="filename">Имя файла</param>
        static void ExecuteApp(string filename)
        {
            if (filename == "") return;
            string arguments = "";
            int pos = filename.IndexOf(" ");
            if (pos > 0)
            {
                arguments = filename.Substring(pos);
                filename = filename.Substring(0, pos);
            }
            if (!filename.Contains(":"))
                filename = System.IO.Directory.GetCurrentDirectory() + "\\" + filename;
            if (!System.IO.File.Exists(filename)) return;
            System.Diagnostics.Process iStartProcess = new System.Diagnostics.Process();
            iStartProcess.StartInfo.FileName = filename;
            iStartProcess.StartInfo.Arguments = arguments;
            //iStartProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            iStartProcess.Start();
            iStartProcess.WaitForExit();
        } // ExecuteApp(string)
    }
}
