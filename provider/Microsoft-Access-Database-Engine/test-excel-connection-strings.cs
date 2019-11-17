using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

class Program {


   private static string[] connectionStringTemplates = new string[] {
// ------------------------------------------------------------------------------------------
                                     @"Data Source=XLSXPATH"                                              , // ArgumentException: An OLE DB Provider was not specified in the ConnectionString.  An example would be, 'Provider=SQLOLEDB;'
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH"                                              , // OleDbException: Unrecognized database format '….xlsx'
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;"                                             , // OleDbException: Unrecognized database format '….xlsx'
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties=Excel 12.0 Xml"           , // OK
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties=Excel 12.0 Xml;HDR=Yes"   , // OleDbException: Could not find installable ISAM.
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties='Excel 12.0 Xml;HDR=Yes'" , // OK
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties='Excel 12.0 Xml'"         , // OK
            @"Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties='Excel 12.0 Xml;HDR=Yes'" , // ArgumentException: An OLE DB Provider was not specified in the ConnectionString.  An example would be, 'Provider=SQLOLEDB;'
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties='Excel 12.0 Xml;IMEX=Yes'", // OK
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties='Excel 12.0 Xml'"         , // OK
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Excel 12.0 Xml"                               , // ArgumentException: Format of the initialization string does not conform to specification
   @"Provider=Microsoft.ACE.OLEDB.12.0;Excel 12.0 Xml;Data Source=XLSXPATH"                               , // OleDbException: Invalid argument
   @"Provider=Microsoft.ACE.OLEDB.12.0;Excel 12.0 Xml;Data Source=XLSXPATH;HDR=Yes"                       , // OleDbException: Invalid argument
   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=XLSXPATH;Extended Properties='Excel 12.0 Xml;foo=1'"   , // OK
   @"Provider=Microsoft.ACE.OLEDB.12.0;Excel 12.0 Xml;Data Source=XLSXPATH;foobarbaz=1"                   , // OK
   @"Provider=Microsoft.ACE.OLEDB.12.0;Excel 12.0 Xml;Data Source=XLSXPATH;foobarbaz"                     , // ArgumentException: Format of the initialization string does not conform to specification
   @"Provider=Microsoft.ACE.OLEDB.12.0;Excel 12.0 Xml;HDR=Yes;Data Source=XLSXPATH;foobarbaz=1"           , // OK
   @"Provider=Microsoft.ACE.OLEDB.12.0;Excel 12.0 Xml;HDR=Yes;Data Source=XLSXPATH"                       , // OK
   @"complete-crap"                                                                                       , // ArgumentException: Format of the initialization string does not conform to specification ….
// ------------------------------------------------------------------------------------------
   };


   static void Main() {

      string xlsxPath = Directory.GetCurrentDirectory() + @"\" + "excelFile.xlsx";

      foreach (string connectionStringTemplate in connectionStringTemplates) {

         string connectionString = connectionStringTemplate.Replace("XLSXPATH", xlsxPath);
         Console.WriteLine(connectionString);

         try {

            using (OleDbConnection connection = new OleDbConnection(connectionString)) {

               Console.WriteLine("  OK create connection");

               OleDbCommand command = new OleDbCommand("select id, num, txt from [tab$]", connection);
               Console.WriteLine("  OK create command ");

               connection.Open();

               OleDbDataReader reader = command.ExecuteReader();
               while (reader.Read()) {
                   Console.WriteLine("\t{0}\t{1}\t{2}", reader[0], reader[1], reader[2]);
               }
               reader.Close();

               Console.WriteLine("  OK test");
            }
         }
         catch (System.ArgumentException argEx) {
            Console.WriteLine($"  NOK: ArgumentException: {argEx.Message}");
         }
         catch (System.Data.OleDb.OleDbException oleDbEx) {
            Console.WriteLine($"  NOK: OleDbException: {oleDbEx.Message}");
         }

         Console.WriteLine("");
      }
   }
}
