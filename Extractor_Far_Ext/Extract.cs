using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;


namespace FAD_EXD_Extractor
{
    public class Extract
    {
        public static string Connect(string filePath)
        {
            OleDbConnectionStringBuilder builder = new OleDbConnectionStringBuilder();
            builder.ConnectionString = @"Data Source=" + filePath;

            // Call the Add method to explicitly add key/value
            // pairs to the internal collection.
            builder.Add("Provider", "Microsoft.Jet.Oledb.4.0");

            return builder.ConnectionString;
            
        }
        
        public static string ReadFAD(string connectionString)
        {
            List<string> sqltableList = new List<string> { @"[Parameters]", @"Evaluation", @"Preliminary" }; // @"Measurement"};

            string createText = null;
            foreach (string queryString in sqltableList)
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {

                    OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    createText = createText + "xxxxxx" + queryString + Environment.NewLine;
                    while (reader.Read())
                    {
                        //Console.WriteLine(reader[0].ToString() + @"," + reader[1].ToString());
                        createText = createText + reader[0].ToString() + @"," + reader[1].ToString() + Environment.NewLine;
                        //Console.WriteLine(createText);
                    }
                    reader.Close();
                }
                //Console.WriteLine(createText);
            }
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
              string queryString = @"Measurement";
              OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

              connection.Open();
              OleDbDataReader reader = command.ExecuteReader();
              createText = createText + "xxxxxx" + queryString + Environment.NewLine;
              while (reader.Read())
              {
                //Console.WriteLine(reader[0].ToString() + @"," + reader[1].ToString());
                createText = createText + reader[0].ToString() + @"," 
                                        + reader[1].ToString() + @","
                                        + reader[2].ToString() + @","
                                        + reader[3].ToString() + @","
                                        + (reader.GetFloat(4)).ToString() + @","
                                        + reader[5].ToString()
                                        + Environment.NewLine;
                //Console.WriteLine(createText);
              }
            reader.Close();
          }
        return createText;
        }
        public static string ReadEXD(string connectionString)
        {
            string[] sqltableList = new string[5];
            sqltableList[0] = @"Parameter";
            sqltableList[1] = @"NumParameter";
            sqltableList[2] = @"Evaluation";
            sqltableList[3] = @"Measurement_1";
            sqltableList[4] = @"Measurement_2";

            string createText = null;
            for(int i = 0; i < 5; i++)
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {

                    OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + sqltableList[i] + ";", connection);

                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    createText = createText + "xxxxxx" + sqltableList[i] + Environment.NewLine;
                    if (i == 2)
                    {
                        while (reader.Read())
                        {
                            createText = createText + (reader[0]).ToString() + @", " 
                                                    + (reader.GetFloat(1)).ToString() + @","
                                                    + (reader.GetFloat(2)).ToString() + @"," 
                                                    + (reader.GetFloat(3)).ToString() + @","
                                                    + (reader.GetFloat(4)).ToString() + @"," 
                                                    + (reader.GetFloat(5)).ToString() + @","
                                                    + (reader.GetFloat(6)).ToString() + @"," 
                                                    + (reader.GetFloat(7)).ToString() + @","
                                                    + (reader.GetFloat(8)).ToString() + @"," 
                                                    + (reader.GetFloat(9)).ToString()
                                                    + Environment.NewLine;
                        }
                        reader.Close();
                    }
                    else if (i==3 | i==4)
                    {
                        while (reader.Read())
                        {
                            createText = createText + (reader.GetFloat(0)).ToString() + @"," 
                                                    + reader[1].ToString() + @","
                                                    + (reader.GetFloat(2)).ToString() + @"," 
                                                    + reader[3].ToString()
                                                    + Environment.NewLine;
                        }
                        reader.Close();
                    }
                    else 
                    {
                        while (reader.Read())
                        {
                            createText = createText + reader[0].ToString() + @"," + reader[1].ToString() + Environment.NewLine;
                        }
                        reader.Close();
                    }      
                }
                
            }
            return createText;
        }
    public static string CsvFAD(string connectionString, int flag)
    {
      List<string> sqltableList = new List<string> { @"[Parameters]", @"Evaluation" }; // , @"Preliminary"@"Measurement"};

      string csvLine = null;
      if (flag == 0) //grab headers
      {
        //csvLine = csvLine + "xxxxxx" + queryString + @",";
        foreach (string queryString in sqltableList)
        {

          using (OleDbConnection connection = new OleDbConnection(connectionString))
          {

            OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();
            csvLine = csvLine + "xxxxxx" + queryString + @",";
            while (reader.Read())
            {
              //Console.WriteLine(reader[0].ToString() + @"," + reader[1].ToString());
              csvLine = csvLine + reader[0].ToString() + @"," ;
              //Console.WriteLine(createText);
            }
            reader.Close();
          }
          //Console.WriteLine(createText);
        }
        csvLine = csvLine + Environment.NewLine;
        /*
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
          string queryString = @"Measurement";
          OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

          connection.Open();
          OleDbDataReader reader = command.ExecuteReader();
          createText = createText + "xxxxxx" + queryString + Environment.NewLine;
          while (reader.Read())
          {
            //Console.WriteLine(reader[0].ToString() + @"," + reader[1].ToString());
            createText = createText + reader[0].ToString() + @","
                                    + reader[1].ToString() + @","
                                    + reader[2].ToString() + @","
                                    + reader[3].ToString() + @","
                                    + (reader.GetFloat(4)).ToString() + @","
                                    + reader[5].ToString()
                                    + Environment.NewLine;
            //Console.WriteLine(createText);
          }
          reader.Close();
        }*/
      } else //grab data
      {
        foreach (string queryString in sqltableList)
        {

          using (OleDbConnection connection = new OleDbConnection(connectionString))
          {

            OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();
            csvLine = csvLine + "xxxxxx" + queryString + @",";
            while (reader.Read())
            {
              //Console.WriteLine(reader[0].ToString() + @"," + reader[1].ToString());
              //createText = createText + reader[0].ToString() + @"," + reader[1].ToString() + Environment.NewLine;
              csvLine = csvLine + reader[1].ToString() + @",";
              //Console.WriteLine(createText);
            }
            reader.Close();
          }
          //Console.WriteLine(createText);
        }
        //csvLine = csvLine + Environment.NewLine;
        /*using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
          string queryString = @"Measurement";
          OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

          connection.Open();
          OleDbDataReader reader = command.ExecuteReader();
          createText = createText + "xxxxxx" + queryString + Environment.NewLine;
          while (reader.Read())
          {
            //Console.WriteLine(reader[0].ToString() + @"," + reader[1].ToString());
            createText = createText + reader[0].ToString() + @","
                                    + reader[1].ToString() + @","
                                    + reader[2].ToString() + @","
                                    + reader[3].ToString() + @","
                                    + (reader.GetFloat(4)).ToString() + @","
                                    + reader[5].ToString()
                                    + Environment.NewLine;
            //Console.WriteLine(createText);
          }
          reader.Close();
        }*/
      }

      return csvLine;
    }
    public static string CsvEXD(string connectionString, int flag)
    {
      string[] sqltableList = new string[5];
      sqltableList[0] = @"Parameter";
      sqltableList[1] = @"NumParameter";
      sqltableList[2] = @"Evaluation";
      //sqltableList[3] = @"Measurement_1";
      //sqltableList[4] = @"Measurement_2";

      string csvLine = null;
      if (flag == 0) //grab headers
      {
        for (int i = 0; i < 3; i++)
        {
          using (OleDbConnection connection = new OleDbConnection(connectionString))
          {

            OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + sqltableList[i] + ";", connection);

            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();

            csvLine = csvLine + "xxxxxx" + sqltableList[i] + @"," ;
            if (i == 2)
            {
              while (reader.Read())
              {
                csvLine = csvLine + //(reader[0]).ToString() + @", "
                                         (reader[0]).ToString() + @" 45min Stretch 1" + @","
                                        + (reader[0]).ToString() + @" 45min Stretch 2" + @","
                                        + (reader[0]).ToString() + @" 90min Stretch 1" + @","
                                        + (reader[0]).ToString() + @" 90min Stretch 2" + @","
                                        + @" " + @","
                                        + @" " + @","
                                        + (reader[0]).ToString() + @" 45min Mean" + @","
                                        + (reader[0]).ToString() + @" 90min Mean" + @","
                                        + @" ";
                                        //+ Environment.NewLine;
              }
              reader.Close();
            }
            else if (i == 3 | i == 4)
            {
              /*while (reader.Read())
              {
                createText = createText + (reader.GetFloat(0)).ToString() + @","
                                        + reader[1].ToString() + @","
                                        + (reader.GetFloat(2)).ToString() + @","
                                        + reader[3].ToString()
                                        + Environment.NewLine;
              }
              reader.Close();*/
            }
            else
            {
              while (reader.Read())
              {
                csvLine = csvLine + reader[0].ToString() + @","; // + reader[1].ToString() + Environment.NewLine;
              }
              reader.Close();
            }
          }
        }
        csvLine = csvLine + Environment.NewLine;
      }
      else //grab data
      {
        for (int i = 0; i < 3; i++)
        {
          using (OleDbConnection connection = new OleDbConnection(connectionString))
          {

            OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + sqltableList[i] + ";", connection);

            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();

            csvLine = csvLine + "xxxxxx" + sqltableList[i] + @","; //+ Environment.NewLine;
            if (i == 2)
            {
              while (reader.Read())
              {
                csvLine = csvLine + //(reader[0]).ToString() + @", "
                                         (reader.GetFloat(1)).ToString() + @","
                                        + (reader.GetFloat(2)).ToString() + @","
                                        + (reader.GetFloat(3)).ToString() + @","
                                        + (reader.GetFloat(4)).ToString() + @","
                                        + (reader.GetFloat(5)).ToString() + @","
                                        + (reader.GetFloat(6)).ToString() + @","
                                        + (reader.GetFloat(7)).ToString() + @","
                                        + (reader.GetFloat(8)).ToString() + @","
                                        + (reader.GetFloat(9)).ToString();
                                        //+ Environment.NewLine;
              }
              reader.Close();
            }
            else if (i == 3 | i == 4)
            {
              while (reader.Read())
              {
                /*createText = createText + (reader.GetFloat(0)).ToString() + @","
                                        + reader[1].ToString() + @","
                                        + (reader.GetFloat(2)).ToString() + @","
                                        + reader[3].ToString()
                                        + Environment.NewLine;*/
              }
              reader.Close();
            }
            else
            {
              while (reader.Read())
              {
                csvLine = csvLine + reader[1].ToString() + @","; // + reader[1].ToString() + Environment.NewLine;
              }
              reader.Close();
            }
          }

        }
      }
      return csvLine;
    }
  }
}
