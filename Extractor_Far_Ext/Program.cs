using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FAD_EXD_Extractor
{
    class Program
    {
        static void Main(string[] args)
        {
             //Arg parsing, help instructions
            if (args.Length == 0)
            {
                        
            string[] fadFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.fad");
            string[] exdFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.exd");

            Console.WriteLine("Found {0} FAD files.", fadFiles.Length);
            Console.WriteLine("Found {0} EXD files.", exdFiles.Length);

            ConsoleKey response;
            do
            {
                Console.Write("Are you sure you want to extract data into text? [y/n] ");
                response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                if (response != ConsoleKey.Enter)
                    Console.WriteLine();

            } while (response != ConsoleKey.Y && response != ConsoleKey.N);

            if(response == ConsoleKey.N)
            { 
            Environment.Exit(0);
            }

            Directory.CreateDirectory("extracted");
            
            foreach (string fadFile in fadFiles)
            {
                string fileName = Path.GetFileName(fadFile);
                Console.Write("Extracting " + fileName);
                string fileContents = Extract.ReadFAD(Extract.Connect(fileName));
                File.WriteAllText((@"extracted\" + fileName.Remove(fileName.Length - 3) + "txt"), fileContents);
                Console.Write("\r");
                Console.WriteLine("Extracting " + fileName + " ...Done");
            }

            foreach (string exdFile in exdFiles)
            {
                string fileName = Path.GetFileName(exdFile);
                Console.Write("Extracting " + fileName);
                string fileContents = Extract.ReadEXD(Extract.Connect(fileName));
                File.WriteAllText((@"extracted\" + fileName.Remove(fileName.Length - 3) + "txt"), fileContents);
                Console.Write("\r");
                Console.WriteLine("Extracting " + fileName + " ...Done");
            }
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();

        } else{
        //there are args
        Console.WriteLine("FAD EXD EXTRACTOR v20201127");
        List<string> fadFiles = new List<string>();
        List<string> exdFiles = new List<string>();
        for (int i = 0; i < args.Length; i++)
        {
          //Console.WriteLine(args[i]);
          //Console.WriteLine(Path.GetFileName(args[i]));
          //Console.WriteLine(Path.GetExtension(args[i]));
          if(Path.GetExtension(args[i]).ToLower() == ".fad")
          {
            fadFiles.Add(Path.GetFileName(args[i]));
          }
          if (Path.GetExtension(args[i]).ToLower() == ".exd")
          {
            exdFiles.Add(Path.GetFileName(args[i]));
          }
        }
        for (int i = 0; i < args.Length; i++)
        {
          Console.WriteLine(args[i]);
        }
        
        Console.WriteLine("Found {0} FAD files.", fadFiles.Count);
        Console.WriteLine("Found {0} EXD files.", exdFiles.Count);
        Console.WriteLine("FAD data will be placed in export/FAD.csv");
        Console.WriteLine("EXD data will be placed in export/EXD.csv");
        ConsoleKey response;
        do
        {
          Console.Write("Are you sure you want to extract data into csv? [y/n] ");
          response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
          if (response != ConsoleKey.Enter)
            Console.WriteLine();

        } while (response != ConsoleKey.Y && response != ConsoleKey.N);

        if (response == ConsoleKey.N)
        {
          Environment.Exit(0);
        }
        Console.WriteLine("Extracting...");
        string appDir = AppDomain.CurrentDomain.BaseDirectory;
        Directory.CreateDirectory(appDir + @"export\");
        if (fadFiles.Count > 0)
        {
          string fileName = fadFiles[0];
          string fileContents = Extract.CsvFAD(Extract.Connect(fileName),0);
          fileContents = fileContents.Insert(0, "Location,");
          File.WriteAllText((appDir + @"export\" + "FAD.csv"), fileContents);

          using (StreamWriter sw = File.AppendText(appDir + @"export\FAD.csv"))
          {
            fileContents = Extract.CsvFAD(Extract.Connect(fileName), 1);
            fileContents = fileContents.Insert(0, (Path.GetFullPath(fileName) + @","));
            sw.WriteLine(fileContents);
            Console.WriteLine(fileName + "...Done");
            for (int i = 1; i < fadFiles.Count; i++)
            {
              fileName = fadFiles[i];
              fileContents = Extract.CsvFAD(Extract.Connect(fileName),1);
              fileContents = fileContents.Insert(0, (Path.GetFullPath(fileName) + @","));
              sw.WriteLine(fileContents);
              Console.WriteLine(fileName + "...Done");
            }
          }
        }
        if (exdFiles.Count > 0)
        {
          string fileName = exdFiles[0];
          string fileContents = Extract.CsvEXD(Extract.Connect(fileName), 0);
          fileContents = fileContents.Insert(0, "Location,");
          File.WriteAllText((appDir + @"export\" + "EXD.csv"), fileContents);

          using (StreamWriter sw = File.AppendText(appDir + @"export\EXD.csv"))
          {
            fileContents = Extract.CsvEXD(Extract.Connect(fileName), 1);
            fileContents = fileContents.Insert(0, (Path.GetFullPath(fileName) + @","));
            //Console.WriteLine(Path.GetFullPath(fileName));
            sw.WriteLine(fileContents);
            Console.WriteLine(fileName + "...Done");
            for (int i = 1; i < exdFiles.Count; i++)
            {
              fileName = exdFiles[i];
              fileContents = Extract.CsvEXD(Extract.Connect(fileName), 1);
              fileContents = fileContents.Insert(0, (Path.GetFullPath(fileName) + @","));
              sw.WriteLine(fileContents);
              Console.WriteLine(fileName + "...Done");
            }
          }
        }

        Console.WriteLine("Press any key to exit");
        Console.ReadKey();
      }
    }

  }
}
