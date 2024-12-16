using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BookcastLexicon
{
    public class Program
    {

        public static SQLiteConnection sqlite_conn;

        static void Main(string[] args)
        {

            sqlite_conn = CreateConnection();

            try 
            {
                CreateTable();
            }

            catch { }

            //DropTable(sqlite_conn);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1()); 
        }

        static SQLiteConnection CreateConnection()
        {

            SQLiteConnection sqlite_conn;
            sqlite_conn = new SQLiteConnection("Data Source=database.db; Version = 3; New = True; Compress = True; ");
         try
            {
                sqlite_conn.Open();
            }
            catch (Exception ex)
            {

            }
            return sqlite_conn;
        }

        static void CreateTable()
        {

            SQLiteCommand sqlite_cmd;
            string Createsql = "CREATE TABLE BookcastDB(Buchtitel TEXT, Folgennummer TEXT, Zeitangabe TEXT, Schlagwort TEXT, Infos TEXT, Quellen TEXT)";
            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = Createsql;
            sqlite_cmd.ExecuteNonQuery();

        }

        public static string CommitValues(string[] values)
        {
            try 
            {
                SQLiteCommand sqlite_cmd;
                sqlite_cmd = sqlite_conn.CreateCommand();
                sqlite_cmd.CommandText = "INSERT INTO BookcastDB (Buchtitel, Folgennummer, Zeitangabe, Schlagwort, Infos, Quellen) " +
                                         "VALUES (@Buchtitel, @Folgennummer, @Zeitangabe, @Schlagwort, @Infos, @Quellen)";

                sqlite_cmd.Parameters.AddWithValue("@Buchtitel", values[0].Trim());
                sqlite_cmd.Parameters.AddWithValue("@Folgennummer", values[1].Trim());
                sqlite_cmd.Parameters.AddWithValue("@Zeitangabe", values[2].Trim());
                sqlite_cmd.Parameters.AddWithValue("@Schlagwort", values[3].Trim());
                sqlite_cmd.Parameters.AddWithValue("@Infos", values[4].Trim());
                sqlite_cmd.Parameters.AddWithValue("@Quellen", values[5].Trim());

                sqlite_cmd.ExecuteNonQuery();
                return "Daten erfolgreich übertragen";
            }

            catch(Exception ex) 
            {
                return ex.Message;
            }


        }

        static void ReadData(SQLiteConnection conn)
        {
            SQLiteDataReader sqlite_datareader;
            SQLiteCommand sqlite_cmd;
            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = "SELECT * FROM SampleTable";

            sqlite_datareader = sqlite_cmd.ExecuteReader();
            while (sqlite_datareader.Read())
            {
                string myreader = sqlite_datareader.GetString(0);
                Console.WriteLine(myreader);
            }
            conn.Close();
        }

        static void DropTable(SQLiteConnection conn) 
        {
                SQLiteCommand sqlite_cmd;

                string dropTableSQL = "DROP TABLE IF EXISTS BookcastDB";

                sqlite_cmd = conn.CreateCommand();
                sqlite_cmd.CommandText = dropTableSQL;
                sqlite_cmd.ExecuteNonQuery();

        }


        public static string FindValue(string value) 
        {

            try
            {

                SQLiteDataReader sqlite_datareader;
                SQLiteCommand sqlite_cmd;
                string output = $"Folgende Informationen sind zum Thema {value} hinterlegt:{Environment.NewLine}{Environment.NewLine}";

                sqlite_cmd = sqlite_conn.CreateCommand();
                sqlite_cmd.CommandText = "SELECT Buchtitel, Folgennummer, Zeitangabe, Infos, Quellen FROM BookcastDB WHERE Schlagwort = @value";
                sqlite_cmd.Parameters.AddWithValue("@value", value);

                sqlite_datareader = sqlite_cmd.ExecuteReader();

                int entryNumber = 1;
                bool found = false; 

                while (sqlite_datareader.Read())
                {
                    found = true;

                    string entry = $"---{entryNumber}---{Environment.NewLine}{Environment.NewLine}";

                    string buchtitel = sqlite_datareader.IsDBNull(0) ? "NULL" : sqlite_datareader["Buchtitel"].ToString();
                    string folgennummer = sqlite_datareader.IsDBNull(1) ? "NULL" : sqlite_datareader["Folgennummer"].ToString();
                    string zeitangabe = sqlite_datareader.IsDBNull(2) ? "NULL" : sqlite_datareader["Zeitangabe"].ToString();
                    string infos = sqlite_datareader.IsDBNull(3) ? "NULL" : sqlite_datareader["Infos"].ToString();
                    string quellen = sqlite_datareader.IsDBNull(4) ? "NULL" : sqlite_datareader["Quellen"].ToString() ;

                    entry += $"{buchtitel}, Folge {folgennummer.Split('#')[1]}, Zeit: {zeitangabe} {Environment.NewLine}{Environment.NewLine}{infos}{Environment.NewLine}{Environment.NewLine}Quellen: {quellen}{Environment.NewLine}{Environment.NewLine}";

                    output += entry;

                    entryNumber++;
                }

                if (!found)
                {
                    return $"Keine Einträge gefunden.";
                }

                return output;
            }

            catch (Exception e)
            {
                return e.Message;
            }
        }


        public static string DeleteValues(string[] values)
        {
            try
            {
                string sqlQuery = "DELETE FROM BookcastDB WHERE Buchtitel = @Buchtitel AND Folgennummer = @Folgennummer AND Zeitangabe = @Zeitangabe AND Schlagwort = @Schlagwort";

                using (SQLiteCommand sqlite_cmd = new SQLiteCommand(sqlQuery, sqlite_conn))
                {
                    sqlite_cmd.Parameters.AddWithValue("@Buchtitel", values[0].Trim());
                    sqlite_cmd.Parameters.AddWithValue("@Folgennummer", values[1].Trim());
                    sqlite_cmd.Parameters.AddWithValue("@Zeitangabe", values[2].Trim());
                    sqlite_cmd.Parameters.AddWithValue("@Schlagwort", values[3].Trim());

                    int rowsAffected = sqlite_cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        return "Die Zeile wurde erfolgreich gelöscht.";
                    }
                    else
                    {
                        return "Keine Übereinstimmung gefunden. Keine Zeile gelöscht.";
                    }
                    }
      
            }
            catch (Exception ex)
            {
                return $"Fehler beim Löschen der Zeile: {ex.Message}";
            }
        }
    }
}