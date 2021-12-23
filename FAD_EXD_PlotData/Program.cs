using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FAD_EXD_PlotData
{
  class Program
  {
    static void Main(string[] args)
    {
      //there are args
      Console.WriteLine("FAD EXD PLOTDATA v20201127");
      List<string> fadFiles = new List<string>();
      List<string> exdFiles = new List<string>();
      for (int i = 0; i < args.Length; i++)
      {
        //Console.WriteLine(args[i]);
        //Console.WriteLine(Path.GetFileName(args[i]));
        //Console.WriteLine(Path.GetExtension(args[i]));
        if (Path.GetExtension(args[i]).ToLower() == ".fad")
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
      Console.WriteLine("FAD data will be placed in export/FADplot.csv");
      Console.WriteLine("EXD data will be placed in export/EXDplot.csv");
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
      //fadFiles.Add(@"C:\Users\Red\Desktop\Work\FarExtExtractorV2\FAD_EXD_PlotData\bin\Debug\farino.FAD");

      if (fadFiles.Count > 0)
      {
        string[,] fadArr = new string[2000, (4 * fadFiles.Count)];
        string fileName = fadFiles[0];
        //string csvLine;
        fadArr[0, 0] = "Location";
        int col = 0;

        for (int i = 0; i < fadFiles.Count; i++)
        {
          fileName = fadFiles[i];
          col = i*4;
          fadArr[0, (col + 1)] = Path.GetFullPath(fileName);
          Extract.CsvFAD(Extract.Connect(fileName), ref fadArr, col);
          Console.WriteLine(fileName + "...Done");
          col = col + 4;
        }
        using (FileStream plotFile = File.Create(appDir + @"export\" + "FADplot.csv"))
        {
        }
          using (StreamWriter sw = File.AppendText(appDir + @"export\FADplot.csv"))
          {
            StringBuilder sb = new StringBuilder(null, 1024);
            for (int j = 0; j < fadArr.GetUpperBound(0); j++)
            {
              for (int i = 0; i < fadArr.GetUpperBound(1); i++)
              {
                sb.Append(fadArr[j, i]+@",");

              }
              //sb.Append(Environment.NewLine);
              sw.WriteLine(sb);
              sb.Clear();
            }
          }
        
          //File.WriteAllText((appDir + @"export\" + "FADplot.csv"), fileContents);
        }
      if (exdFiles.Count > 0)
      {
        string[,] exdArr = new string[600, (8 * exdFiles.Count)];
        string fileName = exdFiles[0];
        exdArr[0, 0] = "Location";
        int col = 0;


        for (int i = 0; i < exdFiles.Count; i++)
        {
          fileName = exdFiles[i];
          col = i * 8;
          exdArr[0, (col + 1)] = Path.GetFullPath(fileName);
          Extract.CsvEXD(Extract.Connect(fileName), ref exdArr, col);
          Console.WriteLine(fileName + "...Done");
          col = col + 8;
        }
        using (FileStream plotFile = File.Create(appDir + @"export\" + "EXDplot.csv"))
        {
        }
        using (StreamWriter sw = File.AppendText(appDir + @"export\EXDplot.csv"))
        {
          StringBuilder sb = new StringBuilder(null, 1024);
          for (int j = 0; j < exdArr.GetUpperBound(0); j++)
          {
            for (int i = 0; i < exdArr.GetUpperBound(1); i++)
            {
              sb.Append(exdArr[j, i] + @",");

            }
            //sb.Append(Environment.NewLine);
            sw.WriteLine(sb);
            sb.Clear();
          }
        }
      }

      Console.WriteLine("Press any key to exit");
      Console.ReadKey();
    }
  }
}
