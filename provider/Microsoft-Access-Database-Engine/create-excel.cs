using System;
using System.Data.OleDb;
using System.IO;

public class CreateExcel {

   static void Main() {
      string excelFilePath = $@"{Directory.GetCurrentDirectory()}\created.xlsx";

      if (File.Exists(excelFilePath))  {
          Console.WriteLine($"{excelFilePath} exists, deleting it.");
          File.Delete(excelFilePath);
      }

      string provider =
              // "Microsoft.Jet.OLEDB.4.0"
                 "Microsoft.ACE.OLEDB.12.0";

      Console.WriteLine($"Using provider {provider}");

      string connetionString = $"Provider={provider};Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=Yes';";

      using (OleDbConnection connection = new OleDbConnection(connetionString)) {

         connection.Open();

         execCommand(connection, "create table tab (id integer, num number, txt varchar)");
         execCommand(connection, "insert into tab values(1, 5, 'five')");
         execCommand(connection, "insert into tab values(2, 2, 'two' )");
         execCommand(connection, "insert into tab values(3, 4, 'four')");
         execCommand(connection, "insert into tab values(4, 9, 'nine')");
      }
   }

   static void execCommand(OleDbConnection conn, string sqlText) {
      using (OleDbCommand command = new OleDbCommand()) {

         command.Connection  = conn;
         command.CommandText = sqlText;
         command.ExecuteNonQuery();

      }
   }
}
