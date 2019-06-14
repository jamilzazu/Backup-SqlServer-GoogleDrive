using Ionic.Zip;
using Ionic.Zlib;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;

namespace BackupSqlServerGoogleDrive
{
    public class Backups
    {
        public static void FullBackup()
        {
            string customer = "MAR";
            string folderPath = @"C:\Program Files\" + customer + @"\";
            string serverName = @"(LocalDb)\MSSQLLocalDB";
            string database = @"Acordos";
            string fileName = string.Format("{0}_{1:yyyyMMdd_HHmmss}.bak", database, DateTime.Now);

            ExistsFolder(folderPath);

            Console.WriteLine("Arquivo: --------------------> " + fileName);

            Server server = new Server(serverName);
            Backup bkpDBFull = new Backup();

            bkpDBFull.Action = BackupActionType.Database;
            bkpDBFull.Database = database;
            bkpDBFull.Devices.AddDevice(folderPath + fileName, DeviceType.File);
            bkpDBFull.BackupSetName = customer + " database Backup";
            bkpDBFull.BackupSetDescription = customer + " database Backup - Full Backup";
            bkpDBFull.ExpirationDate = DateTime.Today.AddDays(10);
            bkpDBFull.Initialize = false;

            bkpDBFull.PercentComplete += CompletionStatusInPercent;
            bkpDBFull.Complete += Backup_Completed;

            bkpDBFull.SqlBackup(server);

            Compress(folderPath, fileName);
        }


        private static void CompletionStatusInPercent(object sender, PercentCompleteEventArgs args)
        {
            Console.Clear();
            Console.WriteLine("Percent completed: {0}%.", args.Percent);
        }

        private static void Backup_Completed(object sender, ServerMessageEventArgs args)
        {
            Console.WriteLine("Backup completed.");
            Console.WriteLine(args.Error.Message);
        }

        private static void Restore_Completed(object sender, ServerMessageEventArgs args)
        {
            Console.WriteLine("Restore completed.");
            Console.WriteLine(args.Error.Message);
        }

        private static void ExistsFolder(string folderPath)
        {
            bool exists = Directory.Exists(folderPath);

            if (!exists)
            {
                Directory.CreateDirectory(folderPath);

                var directoryInfo = new DirectoryInfo(folderPath);
                var directorySecurity = directoryInfo.GetAccessControl();
                var currentUserIdentity = WindowsIdentity.GetCurrent();
                var fileSystemRule = new FileSystemAccessRule(currentUserIdentity.Name,
                                                              FileSystemRights.FullControl,
                                                              InheritanceFlags.ObjectInherit |
                                                              InheritanceFlags.ContainerInherit,
                                                              PropagationFlags.None,
                                                              AccessControlType.Allow);

                directorySecurity.AddAccessRule(fileSystemRule);
                directoryInfo.SetAccessControl(directorySecurity);

                Console.WriteLine("Folder create: ---------------> " + folderPath);
            }
        }

        public static void Compress(string folderPath, string fileName)
        {
            try
            {
                string extension = ".zip";
                string file = folderPath + @"\" + fileName;
                fileName = (fileName + extension).Replace(".bak", "");

                Thread thread = new Thread(t =>
                {
                    using (ZipFile zip = new ZipFile())
                    {
                        zip.Password = "5TCU384VNCH8EVM8";
                        zip.CompressionLevel = CompressionLevel.BestCompression;

                        FileInfo fileInfo = new FileInfo(fileName);
                        zip.AddFile(file, "");
                        DirectoryInfo di = new DirectoryInfo(file);
                        folderPath = string.Format(@"{0}\{1}", di.Parent.FullName, fileInfo.Name);

                        zip.Save(folderPath);

                        File.Delete(file);

                        var shf = new SHFILEOPSTRUCT();
                        shf.wFunc = FO_DELETE;
                        shf.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION;
                        shf.pFrom = file;
                        SHFileOperation(ref shf);

                        Console.WriteLine("Compress completed.");
                    }
                });
                thread.Start();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Compress completed.");
                Console.WriteLine(ex.Message);
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
        public struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.U4)] public int wFunc;
            public string pFrom;
            public string pTo;
            public short fFlags;
            [MarshalAs(UnmanagedType.Bool)] public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern int SHFileOperation(ref SHFILEOPSTRUCT FileOp);

        public const int FO_DELETE = 3;
        public const int FOF_ALLOWUNDO = 0x40;
        public const int FOF_NOCONFIRMATION = 0x10;

    }
}
