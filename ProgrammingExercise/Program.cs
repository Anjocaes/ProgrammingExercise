using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgrammingExercise
{
    internal class Program
    {
        static void Main(string[] args)
        {
            principalMenu();
        }

        private static void principalMenu()
        {
            char option;

            Console.Clear();
            Console.WriteLine("Programming Exercise");
            Console.WriteLine("Menu");
            Console.WriteLine("1. Insert Path");
            Console.WriteLine("2. Exit");
            Console.Write("Option: ");
            option = Convert.ToChar(Console.ReadLine());

            if(option == '1')
                start();
            else
                Console.Write("Thanks for Use");
            Console.ReadLine();
        }

        private static void carga(string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);

            foreach (var fi in di.GetFiles())
            {
                tipoDocumento(fi.FullName, fi.Name);
            }
        }
        private static void start()
        {
            string path;

            Console.Clear();
            Console.Write("Insert Path: ");
            path = @"" + Console.ReadLine();

            if (Directory.Exists(path))
            {
                carga(path);
                startListener(path);
                Console.Clear();
                Console.WriteLine("Waiting...");
                Console.WriteLine("Press enter if you want to finish");
                Console.ReadLine();
                principalMenu();
            }
            else
            {
                Console.WriteLine($"Directory {path} does not exist!");
                Console.WriteLine($"Press Enter to Try Again");
                Console.ReadLine();
                start();
            }
        }

        private static void startListener(string path)
        {
            FileSystemWatcher watcher = new FileSystemWatcher(path);

            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;

            watcher.Created += OnCreated;
            watcher.Error += OnError;

            watcher.EnableRaisingEvents = true;
        }
        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            string value = $"New Document: {e.Name}";
            Console.WriteLine(value);
            if(e.Name != "Master.xls")
                tipoDocumento(e.FullPath,e.Name);
        }
        private static void OnError(object sender, ErrorEventArgs e) =>
            PrintException(e.GetException());

        private static void PrintException(Exception ex)
        {
            if (ex != null)
            {
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine("Stacktrace:");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();
                PrintException(ex.InnerException);
            }
        }

        private static void tipoDocumento(string path, string name)
        {
            bool tipo = name.Contains(".xls");
            int index = path.IndexOf(name);
            string path2 = path.Remove(index, name.Length);

            if (tipo)
                insertarMaster(path, name, path2);
            else
                moverDocumento(path2, name, path, tipo);
        }

        private static void crearMaster(string path)
        {
            if (!File.Exists(path + @"Master.xls"))
            {
                WorkBook master = WorkBook.Create(ExcelFileFormat.XLS);
                WorkSheet xlsSheet = master.CreateWorkSheet("principal");
                master.SaveAs(path + @"Master.xls");
            }
        }
        private static void insertarMaster(string path, string name, string path2)
        {
            crearMaster(path2);

            WorkBook firstBook = WorkBook.Load(path);
            WorkBook secondBook = WorkBook.Load(path2 + @"Master.xls");

            //This is how we can get the first worksheet within the workbook
            WorkSheet worksheet = firstBook.DefaultWorkSheet;

            //This is how we can copy worksheet to the same workbook
            worksheet.CopySheet("Copied Sheet");

            //This is how we can copy worksheet to another workbook with the specified name
            worksheet.CopyTo(secondBook, name);

            secondBook.SaveAs(path2 + @"Master.xls");

            moverDocumento(path2, name, path, true);
        }
        private static void moverDocumento(string path, string name, string ubicacion, bool tipo)
        {
            string finalPathT = path + "Processed" + @"\" + name;
            string finalPathF = path + @"Not Applicable" + @"\"+ name;
            crearCarpeta(path);

            if (tipo)
                File.Move(ubicacion, finalPathT);
            else
                File.Move(ubicacion, finalPathF);

        }
        private static void crearCarpeta(string path)
        {
            string finalPathT = path + "Processed";
            string finalPathF = path + "Not Applicable";

            if (!Directory.Exists(finalPathT))
                Directory.CreateDirectory(path + "Processed");

            if (!Directory.Exists(finalPathF))
                Directory.CreateDirectory(path+ "Not Applicable");
        }

    }
}
