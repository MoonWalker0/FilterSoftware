using ExcelDataReader;
using MahApps.Metro.Controls.Dialogs;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO; 

namespace TelevendFilter
{
    class DatabaseHandling
    {

        public static void CreateTempDB (string path)
        { 
            //Replace xlsx with sqlite.
            SQLiteConnection.CreateFile(path); 

            SQLiteConnection database = new SQLiteConnection("Data Source=" + path + ";Version=3;");
            database.Open();

            string table = "CREATE TABLE CardInfo (" +
                                                    "AccountID VARCHAR(30), " +
                                                    "WorkerID VARCHAR(10), " +
                                                    "StickerID VARCHAR(10), " +
                                                    "Date TIMESTAMP, " +
                                                    "Paid FLOAT, " +
                                                    "Product VARCHAR(30), " +
                                                    "Machine VARCHAR(30))";

            SQLiteCommand command = new SQLiteCommand(table, database);
            command.ExecuteNonQuery();
            database.Close();
             
        }

        public static int GetRowNumber (string path)
        {

            FileStream stream;
            IExcelDataReader reader;
            int output;

            //Verify file validity
            try
            {
                stream = File.Open(path.Remove(path.Length - 6, 6) + "xlsx", FileMode.Open, FileAccess.Read);
            }
            catch(IOException)
            {  
                return -2;
            }            
            
            reader = ExcelReaderFactory.CreateReader(stream);
            output = reader.RowCount;

            //Verify size of file
            try
            {
                reader.Read();
                reader.GetValue(15);
            }
            catch (InvalidOperationException)
            {
                stream.Close();
                stream.Dispose();
                return -1;
            }

            stream.Close();
            stream.Dispose();
            return output;
        }

        public static void FillDB(string path, IProgress<double> controller)
        {
            FileStream stream;
            IExcelDataReader reader;

            stream = File.Open(path.Remove(path.Length - 6, 6) + "xlsx", FileMode.Open, FileAccess.Read);
            reader = ExcelReaderFactory.CreateReader(stream);

            SQLiteConnection database;
            SQLiteCommand command;
            int currentRow;
            string format;
            string transactionType;
            string Paid = ""; 

            database = new SQLiteConnection("Data Source=" + path + ";Version=3;");
            database.Open(); 
            
            currentRow = 2;
            //Skip first row
            reader.Read();

            while (reader.Read())
            {
                transactionType = Convert.ToString(reader.GetValue(8));

                if (transactionType == null)
                {
                    break;
                }
                else if (Convert.ToString(reader.GetValue(5)).StartsWith("leg"))
                {
                    //Do nothing - ignore customer events
                    currentRow++;
                    continue;
                }
                else if (transactionType == "PaymentCashless")
                {
                    Paid = Convert.ToString(reader.GetValue(12)).Replace(",",".");  
                }
                else if (transactionType == "DebtBuyout" ||
                         transactionType == "BusinessBonus" ||
                         transactionType == "RechargeWebRefund")
                {
                    //Ignore reload events
                    currentRow++;
                    continue;
                }


                format = "INSERT INTO CardInfo " +
                         "(AccountID, Date, Paid, Product, Machine) values " +
                         "(\"" + Convert.ToString(reader.GetValue(7)) + "\", " +
                         "\"" + Convert.ToString(reader.GetValue(0)) + "\", " +
                         Paid + ", " +
                         "\"" + Convert.ToString(reader.GetValue(10)) + "\"," +
                         "\"" + Convert.ToString(reader.GetValue(14)) + "\")";


                command = new SQLiteCommand(format, database);
                command.ExecuteNonQuery();
                currentRow++;
                controller.Report(Convert.ToDouble(currentRow - 2));
            }
             
            stream.Close();
            stream.Dispose();
            database.Close(); 
        }

        public static int AddWorkerID(string pathDB, string pathFile)
        {
            FileStream stream;
            IExcelDataReader reader;
            SQLiteConnection database;
            SQLiteCommand command;
            int currentRow;
            string format;
            string targetID;

            try
            { 
                stream = File.Open(pathFile, FileMode.Open, FileAccess.Read);
            }
            catch (IOException)
            {
                return 0;
            }
            reader = ExcelReaderFactory.CreateReader(stream);

            database = new SQLiteConnection("Data Source=" + pathDB + ";Version=3;");
            database.Open(); 

            currentRow = 3;
            //Skip to the third row
            reader.Read();
            reader.Read();
            while (reader.Read())
            {
                targetID = Convert.ToString(reader.GetValue(3));

                if (targetID == null)
                {
                    break;
                }

                format = "UPDATE CardInfo " +
                         "SET WorkerID = " +
                         "\"" + Convert.ToString(reader.GetValue(1)) + "\", " +
                         "StickerID = " +
                         "\"" + Convert.ToString(reader.GetValue(2)) + "\" " +
                         "WHERE AccountID = \"" + targetID + "\"";

                command = new SQLiteCommand(format, database);
                command.ExecuteNonQuery();
                currentRow++;
            }

            stream.Close();
            stream.Dispose();
            database.Close();

            return 1;
        }

        public static List<ListDisplay.ListItem> LoadAllByDate(string path)
        {
            string commandString = "SELECT * FROM CardInfo ORDER BY date";
            SQLiteConnection database = new SQLiteConnection("Data Source=" + path + ";Version=3;");
            SQLiteCommand command = new SQLiteCommand(commandString, database);
            List<ListDisplay.ListItem> output = new List<ListDisplay.ListItem>();
            SQLiteDataReader rowData;

            database.Open();
            rowData = command.ExecuteReader();

            /*Populate collection*/
            while (rowData.Read())
            {
                output.Add(ListDisplay.FormListItem(rowData));
            }

            database.Close();

            return output;
        }

        public static List<ListDisplay.ListItem> PerformFilter(DateTime? fromDate, DateTime? toDate,
                                                               string filterCard, string ignoreCard, string path)
        {
            string commandString;
            SQLiteConnection database = new SQLiteConnection("Data Source=" + path + ";Version=3;");
            SQLiteCommand command;
            List<ListDisplay.ListItem> output = new List<ListDisplay.ListItem>();
            SQLiteDataReader rowData;
            bool flagDate = false, flagOne = false;

            //Form command
            commandString = "SELECT * FROM CardInfo WHERE ";

            //Date filtering
            if (fromDate != null && toDate != null)
            {
                commandString += "Date BETWEEN \"" + Utilities.FormSQLDateFormat(fromDate.Value, true)
                              + "\" AND \"" + Utilities.FormSQLDateFormat(toDate.Value, false) + "\" ";
                flagDate = true;
            } 
            else if (fromDate == null && toDate != null)
            {
                commandString += "Date < \"" + Utilities.FormSQLDateFormat(toDate.Value, true) + "\" ";
                flagDate = true;
            }
            else if (toDate == null && fromDate != null)
            {
                commandString += "Date > \"" + Utilities.FormSQLDateFormat(fromDate.Value, false) + "\" ";
                flagDate = true;
            }

            //Special snowflake
            if (filterCard != "Kod karty" && filterCard != null && filterCard != "")
            {
                if (flagDate)
                {
                    commandString += "AND ";
                }
                commandString += "AccountID = \"" + filterCard + "\" ";
                flagOne = true;
            }

            //Ignore card 
            if (ignoreCard != "Kod karty" && ignoreCard != null && ignoreCard != "")
            {
                if(flagDate || flagOne)
                {
                    commandString += "AND ";
                }
                commandString += "AccountID != \"" + ignoreCard + "\" ";
            }

            commandString += "ORDER BY date";

            command = new SQLiteCommand(commandString, database);

            database.Open();
            try
            {
                rowData = command.ExecuteReader();
            }
            catch(System.Data.SQLite.SQLiteException)
            {
                return null;
            }

            /*Populate collection*/
            while (rowData.Read())
            {
                output.Add(ListDisplay.FormListItem(rowData));
            }

            database.Close();

            return output;

        }

        public static List<string> GetCardIDs(string path)
        {
            string commandString = "SELECT DISTINCT AccountID FROM CardInfo";
            SQLiteConnection database = new SQLiteConnection("Data Source=" + path + ";Version=3;");
            SQLiteCommand command = new SQLiteCommand(commandString, database);
            List<string> output = new List<string>();
            SQLiteDataReader returned;

            database.Open();
            returned = command.ExecuteReader();

            /*Populate collection*/
            while (returned.Read())
            {
                output.Add(returned["AccountID"].ToString());
            }

            database.Close();

            return output;
        }
    }
}
