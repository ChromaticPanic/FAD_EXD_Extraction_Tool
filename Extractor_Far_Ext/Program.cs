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
            /* Arg parsing, help instructions
            if (args.Length == 0)
            {
                Console.WriteLine("Please enter a numeric argument.");
                Console.WriteLine("Usage: Factorial <num>");
                return 1;
            }*/

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

        }
       
    }
}
