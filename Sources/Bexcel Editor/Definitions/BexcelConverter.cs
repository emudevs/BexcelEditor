using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace Bexcel_Editor.Definitions
{
    internal static class BexcelConverter
    {
        static readonly string[] SQLiteIndexes = {"CREATE INDEX '1' ON '1_LongSword' ('Index')"
                , "CREATE INDEX '2' ON '2_Blunt' ('Index')"
                , "CREATE INDEX '3' ON '3_TwoHandSword' ('Index')"
                , "CREATE INDEX '4' ON '4_Bow' ('Index')"
                , "CREATE INDEX '5' ON '5_Dagger' ('Index')"
                , "CREATE INDEX '6' ON 'character_table' ('Index')"
                , "CREATE INDEX '7' ON 'item_table' ('Index')"
                , "CREATE INDEX '8' ON 'buff_table' ('Index')"
        };

        public static Bexcel Read(string inputPath)
        {
            try
            {
                var bexcel = new Bexcel();

                using (BinaryReader bReader = new BinaryReader(File.Open(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read), Encoding.Unicode))
                {
                    int sheetCount = bReader.ReadInt32();

                    for (int i = 0; i < sheetCount; i++)
                    {
                        var tableName = ReadString(bReader);
                        var tableType = bReader.ReadInt32();

                        bexcel.Sheets.Add(new Bexcel.Sheet
                        {
                            Name = tableName,
                            Type = tableType
                        });
                    }

                    sheetCount = bReader.ReadInt32();
                    for (int j = 0; j < sheetCount; j++)
                    {
                        string sheetName = ReadString(bReader);
                        var currentSheet = bexcel.Sheets.First(x => x.Name == sheetName);

                        currentSheet.Index1 = bReader.ReadInt32();
                        currentSheet.Index2 = bReader.ReadInt32();

                        int columnCount = bReader.ReadInt32();
                        for (int k = 0; k < columnCount; k++)
                        {
                            currentSheet.Columns.Add(new Bexcel.Column
                            {
                                //Index = k,
                                Name = ReadString(bReader)
                            });
                        }

                        int rowCount = bReader.ReadInt32();
                        for (int l = 0; l < rowCount; l++)
                        {
                            Bexcel.Row row = new Bexcel.Row
                            {
                                //Index = l,
                                Index1 = bReader.ReadInt32(),
                                Index2 = bReader.ReadInt32()
                            };

                            int cellCount = bReader.ReadInt32();

                            for (int m = 0; m < cellCount; m++)
                            {
                                row.Cells.Add(new Bexcel.Cell
                                {
                                    Index = m,
                                    Name = ReadString(bReader)
                                });
                            }

                            currentSheet.Rows.Add(row);
                        }

                        int columns = bReader.ReadInt32();
                        for (int n = 0; n < columns; n++)
                        {
                            currentSheet.Unknown1.Add(new Bexcel.Unknown
                            {
                                Index = n,
                                Text = ReadString(bReader),
                                Number = bReader.ReadInt32()
                            });
                        }

                        int rowCount2 = bReader.ReadInt32();
                        for (int num8 = 0; num8 < rowCount2; num8++)
                        {
                            currentSheet.TableDetails.Add(new Bexcel.Unknown
                            {
                                Index = num8,
                                Text = ReadString(bReader),
                                Number = bReader.ReadInt32()
                            });
                        }
                    }

                    bexcel.FileEnding = ReadString(bReader);
                }

                return bexcel;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error happened while reading Bexcel file. Is the file corrupted or not compatible?!\n\n{ex}");
            }

            return null;
        }

        public static void Save(Bexcel bexcel, string outputPath)
        {
            // TODO: Test Index1 and Index2 values to see if they are important.
            try
            {
                // If the output file exists, remove the file
                if (File.Exists(outputPath))
                    File.Delete(outputPath);

                // Using a file stream, to be automatically disposed afterwards
                using (var bw = new BinaryWriter(new FileStream(outputPath, FileMode.CreateNew), Encoding.Unicode))
                {
                    // Table count
                    bw.Write(bexcel.Sheets.Count);

                    // Foreach table
                    foreach (var sheet in bexcel.Sheets.OrderBy(x => x.Type))
                    {
                        // Table name
                        bw.Write((long)sheet.Name.Length);
                        bw.Write(WriteString(sheet.Name));

                        // Table type (unknown?)
                        bw.Write(sheet.Type);
                    }

                    // Table count
                    bw.Write(bexcel.Sheets.Count); // sheetCount - Int32

                    // Foreach table in bexcel
                    foreach (var sheet in bexcel.Sheets.OrderBy(x => x.Type))
                    {
                        // Write table name
                        bw.Write((long)sheet.Name.Length);
                        bw.Write(WriteString(sheet.Name));

                        // Write Index1 value from memory
                        bw.Write(sheet.Index1);
                        // Write Index2 value from memory
                        bw.Write(sheet.Index2);

                        // Write column count
                        bw.Write(sheet.Columns.Count);
                        foreach (var column in sheet.Columns)
                        {
                            // Write column name
                            bw.Write((long)column.Name.Length);
                            bw.Write(WriteString(column.Name));
                        }

                        // Write row count
                        bw.Write(sheet.Rows.Count);
                        foreach (var row in sheet.Rows)
                        {
                            // Write Index1 value from memory
                            bw.Write(row.Index1);

                            // Write Index2 value from memory
                            bw.Write(row.Index2);

                            // Write cell count
                            bw.Write(row.Cells.Count);

                            // Foreach cells
                            foreach (var cells in row.Cells)
                            {
                                // Write cell value
                                bw.Write((long)cells.Name.Length);
                                bw.Write(WriteString(cells.Name));
                            }
                        }

                        // Contains column names without unused or null ones, TODO: Is it necessary?
                        bw.Write(sheet.Unknown1.Count);
                        foreach (var unk in sheet.Unknown1)
                        {
                            bw.Write((long)unk.Text.Length);
                            bw.Write(WriteString(unk.Text));
                            bw.Write(unk.Number);
                        }

                        // Contains table details such as keys, TODO: Is it necessary?
                        bw.Write(sheet.TableDetails.Count);
                        foreach (var unk in sheet.TableDetails)
                        {
                            bw.Write((long)unk.Text.Length);
                            bw.Write(WriteString(unk.Text));
                            bw.Write(unk.Number);
                        }
                    }

                    // TODO: Is it necessary?
                    bw.Write((long)bexcel.FileEnding.Length);
                    bw.Write(WriteString(bexcel.FileEnding));
                }

                MessageBox.Show("Bexcel file created at: " + outputPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error happened while generating Bexcel file from Datasheets in memory.\n\n{ex}");
            }
        }

        public static void SaveAsSQLite(Bexcel bexcel, string outputPath)
        {
            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);

                string cs = $"Data Source={outputPath};Version=3;";
                using (SQLiteConnection sqlCon = new SQLiteConnection(cs))
                {
                    Debug.WriteLine("Generating SQLite3 Database");
                    sqlCon.Open();

                    foreach (var sheet in bexcel.Sheets.OrderBy(x => x.Name))
                    {
                        List<Tuple<string, int>> duplicates = new List<Tuple<string, int>>();

                        string createString = $"create table `{sheet.Name}` (ID INTEGER PRIMARY KEY AUTOINCREMENT";

                        var nullColumnCount = sheet.Columns.Count(x => x.Name == "<null>");
                        var currentNullColumn = 1;

                        Debug.WriteLine("Creating table: " + sheet.Name);
                        using (var cmd = new SQLiteCommand(sqlCon))
                        {
                            using (var transaction = sqlCon.BeginTransaction())
                            {
                                for (int i = 0; i < sheet.Columns.Count; i++)
                                {
                                    Bexcel.Column column = sheet.Columns[i];
                                    string extraString;

                                    if (column.Name == "<null>")
                                    {
                                        extraString = $"<null>{currentNullColumn++}";
                                    }
                                    else if (sheet.Columns.Count(x => x.Name == column.Name) > 1)
                                    {
                                        int rowCount = sheet.Rows.Count(x => x.Cells[i].Name != "<null>");

                                        extraString = (rowCount == 0) ? $"{column.Name}_Null" : (duplicates.Count(x => x.Item1 == column.Name && x.Item2 == rowCount) > 0) ? $"{column.Name}_Null" : column.Name;

                                        if (duplicates.Count(x => x.Item1 == column.Name && x.Item2 == rowCount) == 0)
                                        {
                                            duplicates.Add(new Tuple<string, int>(column.Name, rowCount));
                                        }

                                        Debug.WriteLine($"[{sheet.Name}] Found a duplicate column name: {column.Name}");
                                    }
                                    else
                                    {
                                        extraString = column.Name;
                                    }

                                    createString += ", `" + extraString + "` NTEXT";
                                }
                                createString += ")";

                                cmd.CommandText = createString;
                                cmd.ExecuteNonQuery();

                                transaction.Commit();
                            }
                        }

                        using (var cmd = new SQLiteCommand(sqlCon))
                        {
                            using (var transaction = sqlCon.BeginTransaction())
                            {
                                foreach (var sheetRows in sheet.Rows)
                                {
                                    string InsertSql = $"INSERT INTO `{sheet.Name}` values (NULL";
                                    for (int i = 0; i < sheetRows.Cells.Count; i++)
                                    {
                                        InsertSql += (sheetRows.Cells[i].Name == "<null>") ? ", NULL" : $", '{sheetRows.Cells[i].Name.Replace("'", "''")}'";
                                    }

                                    InsertSql += ")";
                                    cmd.CommandText = InsertSql;
                                    cmd.ExecuteNonQuery();
                                }

                                transaction.Commit();
                            }
                        }
                    }

                    // Create table Indexes
                    using (var cmd = new SQLiteCommand(sqlCon))
                    {
                        using (var transaction = sqlCon.BeginTransaction())
                        {
                            for (int i = 0; i < SQLiteIndexes.Length; i++)
                            {
                                cmd.CommandText = SQLiteIndexes[i];
                                cmd.ExecuteNonQuery();
                            }

                            transaction.Commit();
                        }
                    }

                    sqlCon.Close();

                    MessageBox.Show("SQLite file created at: " + outputPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error happened while generating SQLite file.\n\n{ex}");
            }
        }

        public static void Dump<T>(this T x)
        {
            Debug.WriteLine(JsonConvert.SerializeObject(x, Formatting.Indented, new JsonSerializerSettings
            {
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            }));
        }

        public static DataTable ToDataTable(Bexcel.Sheet sheet)
        {
            try
            {
                DataTable dt = new DataTable();

                foreach (var column in sheet.Columns)
                {
                    dt.Columns.Add(column.Name);
                }

                foreach (var row in sheet.Rows)
                {
                    var dr = dt.NewRow();
                    int i = 0;
                    foreach (var cell in row.Cells)
                    {
                        dr[i] = cell.Name;
                        i++;
                    }
                    dt.Rows.Add(dr);
                }

                dt.TableName = sheet.Name;

                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error happened while generating DataTable from Bexcel sheet.\n\n{ex}");
                return new DataTable();
            }
        }

        private static string ReadString(BinaryReader r)
        {
            return Encoding.Unicode.GetString(r.ReadBytes((int)r.ReadInt64() * 2));
        }

        private static byte[] WriteString(string s)
        {
            return Encoding.Unicode.GetBytes(s);
        }
    }
}