using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace BookcastLexicon
{
    public class Program
    {

        public static SQLiteConnection sqlite_conn;

        [STAThread]
        static void Main(string[] args)
        {

            sqlite_conn = CreateConnection();

            try 
            {
                CreateTable();
            }

            catch { }

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

        public static void CreateTable()
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

        public static void DropTable()
        {
            DialogResult dialogResult = MessageBox.Show("Möchten Sie die Datenbank wirklich löschen?", "Bestätigung", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    string dropTableSQL = "DROP TABLE IF EXISTS BookcastDB";

                    SQLiteCommand sqlite_cmd = sqlite_conn.CreateCommand();
                    sqlite_cmd.CommandText = dropTableSQL;
                    sqlite_cmd.ExecuteNonQuery();

                    MessageBox.Show("Die Datenbanktabelle wurde erfolgreich gelöscht.", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fehler beim Löschen der Datenbanktabelle: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Das Löschen der Tabelle wurde abgebrochen.", "Abgebrochen", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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

        public static void ExportDatabaseToExcel()
        {
            try
            {
                string query = "SELECT * FROM BookcastDB";

                SQLiteCommand sqlite_cmd = sqlite_conn.CreateCommand();
                sqlite_cmd.CommandText = query;

                SQLiteDataReader reader = sqlite_cmd.ExecuteReader();

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("BookcastDB");

                    worksheet.Cell(1, 1).Value = "Buchtitel";
                    worksheet.Cell(1, 2).Value = "Folgennummer";
                    worksheet.Cell(1, 3).Value = "Zeitangabe";
                    worksheet.Cell(1, 4).Value = "Schlagwort";
                    worksheet.Cell(1, 5).Value = "Infos";
                    worksheet.Cell(1, 6).Value = "Quellen";

                    int row = 2; 

                    while (reader.Read())
                    {
                        worksheet.Cell(row, 1).Value = reader.IsDBNull(0) ? "NULL" : reader["Buchtitel"].ToString();
                        worksheet.Cell(row, 2).Value = reader.IsDBNull(1) ? "NULL" : reader["Folgennummer"].ToString();
                        worksheet.Cell(row, 3).Value = reader.IsDBNull(2) ? "NULL" : reader["Zeitangabe"].ToString();
                        worksheet.Cell(row, 4).Value = reader.IsDBNull(3) ? "NULL" : reader["Schlagwort"].ToString();
                        worksheet.Cell(row, 5).Value = reader.IsDBNull(4) ? "NULL" : reader["Infos"].ToString();
                        worksheet.Cell(row, 6).Value = reader.IsDBNull(5) ? "NULL" : reader["Quellen"].ToString();

                        row++;
                    }

                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Excel-Dateien (*.xlsx)|*.xlsx"; 
                        saveFileDialog.DefaultExt = ".xlsx";  
                        saveFileDialog.FileName = "BookcastDB_Export.xlsx";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string filePath = saveFileDialog.FileName;
                            workbook.SaveAs(filePath);
                            MessageBox.Show("Daten wurden erfolgreich in eine Excel-Datei exportiert.", "Export Erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Exportieren der Daten: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public static void ImportDataFromExcel()
        {
            DialogResult dialogResult = MessageBox.Show("Möchten Sie die bestehende Datenbank überschreiben?", "Bestätigung", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (dialogResult == DialogResult.Yes)
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel-Dateien (*.xls; *.xlsx)|*.xls; *.xlsx"; 
                    openFileDialog.Title = "Wählen Sie eine Excel-Datei zum Importieren aus";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;

                        if (File.Exists(filePath))
                        {
                            try
                            {
                                Excel.Application xlApp = new Excel.Application();
                                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                                Excel.Range xlRange = xlWorksheet.UsedRange;

                                string dropTableSQL = "DROP TABLE IF EXISTS BookcastDB";
                                SQLiteCommand sqlite_cmd = sqlite_conn.CreateCommand();
                                sqlite_cmd.CommandText = dropTableSQL;
                                sqlite_cmd.ExecuteNonQuery();

                                CreateTable();

                                for (int row = 2; row <= xlRange.Rows.Count; row++) 
                                {
                                    string buchtitel = xlRange.Cells[row, 1].Text.ToString();
                                    string folgennummer = xlRange.Cells[row, 2].Text.ToString();
                                    string zeitangabe = xlRange.Cells[row, 3].Text.ToString();
                                    string schlagwort = xlRange.Cells[row, 4].Text.ToString();
                                    string infos = xlRange.Cells[row, 5].Text.ToString();
                                    string quellen = xlRange.Cells[row, 6].Text.ToString();

                                    string insertDataSQL = "INSERT INTO BookcastDB (Buchtitel, Folgennummer, Zeitangabe, Schlagwort, Infos, Quellen) " +
                                                           "VALUES (@buchtitel, @folgennummer, @zeitangabe, @schlagwort, @infos, @quellen)";

                                    sqlite_cmd.CommandText = insertDataSQL;
                                    sqlite_cmd.Parameters.AddWithValue("@buchtitel", buchtitel);
                                    sqlite_cmd.Parameters.AddWithValue("@folgennummer", folgennummer);
                                    sqlite_cmd.Parameters.AddWithValue("@zeitangabe", zeitangabe);
                                    sqlite_cmd.Parameters.AddWithValue("@schlagwort", schlagwort);
                                    sqlite_cmd.Parameters.AddWithValue("@infos", infos);
                                    sqlite_cmd.Parameters.AddWithValue("@quellen", quellen);

                                    sqlite_cmd.ExecuteNonQuery();
                                }

                                xlWorkbook.Close();
                                xlApp.Quit();

                                MessageBox.Show("Daten wurden erfolgreich importiert und die Datenbank überschrieben.",
                                                "Import Erfolgreich",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Fehler beim Importieren der Excel-Daten: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Die Datenbanküberschreibung wurde abgebrochen.", "Abgebrochen", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}