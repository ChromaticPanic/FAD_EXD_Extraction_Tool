using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;


namespace FAD_EXD_PlotData
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
        
    public static bool CsvFAD(string connectionString, ref string[,] fadArr, int col)
    {
      List<string> sqltableList = new List<string> { @"[Parameters]", @"Evaluation" }; // , @"Preliminary"@"Measurement"};
      int j = col;
      
      if (col == 0)
      {
        foreach (string queryString in sqltableList)
        {

          using (OleDbConnection connection = new OleDbConnection(connectionString))
          {

            OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();
            int i = 1;
            while (reader.Read())
            {
              fadArr[i, j] = reader[0].ToString();
              fadArr[i, j+1] = reader[1].ToString();
              i++;
            }
            reader.Close();
          }
          j = j + 2;
        }
      }
      else
      {
        foreach (string queryString in sqltableList)
        {

          using (OleDbConnection connection = new OleDbConnection(connectionString))
          {

            OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();
            int i = 1;
            while (reader.Read())
            {
              //fadArr[i, col] = reader[0].ToString();
              fadArr[i, j + 1] = reader[1].ToString();
              i++;
            }
            reader.Close();
          }
          j = j + 2;
        }
      }
        
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
          string queryString = @"Measurement";
          OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + queryString + ";", connection);

          connection.Open();
          OleDbDataReader reader = command.ExecuteReader();
          int i = 30;
          while (reader.Read())
          {
          fadArr[i, col] = reader[0].ToString();
          fadArr[i, col+1] = reader[1].ToString();
          fadArr[i, col+2] = reader[2].ToString();
          fadArr[i, col+3] = reader[3].ToString();
          i++;
          }
          reader.Close();
        }
      

      return true;
    }
    public static bool CsvEXD(string connectionString, ref string[,] exdArr, int col)
    {
      //string[] sqltableList = new string[2];
      //sqltableList[0] = @"Parameter";
      //sqltableList[1] = @"NumParameter";
      //sqltableList[2] = @"Evaluation";
      //sqltableList[0] = @"Measurement_1";
      //sqltableList[1] = @"Measurement_2";

      int j = col;
      int i = 1;

      using (OleDbConnection connection = new OleDbConnection(connectionString))
      {

        OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + @"Parameter" + ";", connection);

        connection.Open();
        OleDbDataReader reader = command.ExecuteReader();
        
        while (reader.Read())
        {
          exdArr[i, j] = reader[0].ToString();
          exdArr[i, j + 1] = reader[1].ToString();
          i++;
        }
        reader.Close();
      }

      j = col;

      using (OleDbConnection connection = new OleDbConnection(connectionString))
      {

        OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + @"Evaluation" + ";", connection);

        connection.Open();
        OleDbDataReader reader = command.ExecuteReader();

        exdArr[i, j] = " ";
        exdArr[i, j + 1] = "45min Rep 1";
        exdArr[i, j + 2] = "45min Rep 2";
        exdArr[i, j + 3] = "90min Rep 1";
        exdArr[i, j + 4] = "90min Rep 1";
        //(reader.GetFloat(5)).ToString();
        //(reader.GetFloat(6)).ToString();
        exdArr[i, j + 5] = "45min Mean";
        exdArr[i, j + 6] = "90min Mean";
        //(reader.GetFloat(9)).ToString();
        i++;

        while (reader.Read())
        {
          exdArr[i, j] = reader[0].ToString();
          exdArr[i, j+1] = (reader.GetFloat(1)).ToString();
          exdArr[i, j+2] = (reader.GetFloat(2)).ToString();
          exdArr[i, j+3] = (reader.GetFloat(3)).ToString();
          exdArr[i, j+4] = (reader.GetFloat(4)).ToString();
          //(reader.GetFloat(5)).ToString();
          //(reader.GetFloat(6)).ToString();
          exdArr[i, j+5] = (reader.GetFloat(7)).ToString();
          exdArr[i, j+6] = (reader.GetFloat(8)).ToString();
          //(reader.GetFloat(9)).ToString();
          i++;
        }
        reader.Close();
      }
      
      exdArr[i, j + 0] = "45min Rep 1 mm";
      exdArr[i, j + 1] = "45min Rep 1 BU";
      exdArr[i, j + 2] = "45min Rep 2 mm";
      exdArr[i, j + 3] = "45min Rep 2 BU";
      exdArr[i, j + 4] = "90min Rep 1 mm";
      exdArr[i, j + 5] = "90min Rep 1 BU";
      exdArr[i, j + 6] = "90min Rep 2 mm";
      exdArr[i, j + 7] = "90min Rep 2 BU";
      i++;
      int k = i;

      using (OleDbConnection connection = new OleDbConnection(connectionString))
      {

        OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + @"Measurement_1" + ";", connection);

        connection.Open();
        OleDbDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {
          exdArr[i, j] = (reader.GetFloat(0)).ToString();
          exdArr[i, j + 1] = reader[1].ToString();
          exdArr[i, j + 2] = (reader.GetFloat(2)).ToString();
          exdArr[i, j + 3] = reader[3].ToString();
          i++;
        }
        reader.Close();
        j = j + 4;
      }
      using (OleDbConnection connection = new OleDbConnection(connectionString))
      {
        OleDbCommand command = new OleDbCommand(@"SELECT * FROM " + @"Measurement_2" + ";", connection);

        connection.Open();
        OleDbDataReader reader = command.ExecuteReader();

        while (reader.Read())
        {
          exdArr[k, j] = (reader.GetFloat(0)).ToString();
          exdArr[k, j + 1] = reader[1].ToString();
          exdArr[k, j + 2] = (reader.GetFloat(2)).ToString();
          exdArr[k, j + 3] = reader[3].ToString();
          k++;
        }
        reader.Close();
      }
      
      return true;
    }
  }
}
