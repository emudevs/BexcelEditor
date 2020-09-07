using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;

namespace Bexcel_Editor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        internal static Definitions.Bexcel Bexcel;

        public MainWindow()
        {
            InitializeComponent();

            //OpenBexcelFile(@"D:\Projeler\BDO\Bexcel-Editor\Tests\datasheets.bexcel");
        }

        private void BexcelFileDrop(object sender, DragEventArgs e)
        {
            try
            {
                string[] DroppedItems = (string[])e.Data.GetData(DataFormats.FileDrop, false);

                if (DroppedItems == null)
                {
                    return;
                }

                foreach (string DroppedItem in DroppedItems)
                {
                    var Info = new FileInfo(DroppedItem);

                    if ((Info.Attributes & FileAttributes.Archive) != 0 && (Info.Extension == ".bexcel"))
                    {
                        OpenBexcelFile(Info.FullName);
                        break; // 
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error happened with file drop function.\n\n" + ex);
            }
        }

        private void OpenBexcelFile(string fileLocation)
        {
            try
            {
                Bexcel = Definitions.BexcelConverter.Read(fileLocation);

                UpdateItemSource();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error happened with Open Bexcel function.\n\n" + ex);
            }
        }

        private void UpdateItemSource(string Search = "")
        {
            Sheets.ItemsSource = (string.IsNullOrEmpty(Search) ? Bexcel.Sheets.OrderBy(x => x.Name) : Bexcel.Sheets.Where(x => x.Name.IndexOf(Search, StringComparison.InvariantCultureIgnoreCase) >= 0).OrderBy(x => x.Name));
        }

        private void Border_DragEnter(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
        }

        private void SaveBexcelFile_Click(object sender, RoutedEventArgs e)
        {
            if (Bexcel == null)
                return;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".bexcel",
                Filter = "Bexcel file (*.bexcel)|*.bexcel"
            };

            if (dlg.ShowDialog() == true)
            {
                Definitions.BexcelConverter.Save(Bexcel, dlg.FileName);
            }
        }

        private void SaveSQLiteFile_Click(object sender, RoutedEventArgs e)
        {
            if (Bexcel == null)
                return;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".sqlite3",
                Filter = "SQLite file (*.sqlite3)|*.sqlite3"
            };

            if (dlg.ShowDialog() == true)
            {
                Definitions.BexcelConverter.SaveAsSQLite(Bexcel, dlg.FileName);
            }
        }

        private void Sheets_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (!(Sheets.SelectedItem is Definitions.Bexcel.Sheet sheet))
                    return;

                Sheet.ItemsSource = sheet.Dt.DefaultView;

                sheet.Dt.RowChanged += Dt_RowChanged;
                sheet.Dt.RowDeleting += Dt_RowDeleting;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error happened with updating the sheet selection.\n\n" + ex);
            }
        }

        private void Dt_RowChanged(object sender, DataRowChangeEventArgs e)
        { // Change & Add
            try
            {
                var sheet = Bexcel.Sheets.First(x => x.Name == ((dynamic)sender).TableName);

                if (e.Action == DataRowAction.Add)
                {
                    var bexcelRow = new Definitions.Bexcel.Row
                    {
                        Index1 = -1,
                        Index2 = 1
                    };

                    for (int i = 0; i < e.Row.ItemArray.Length; i++)
                    {
                        bexcelRow.Cells.Add(new Definitions.Bexcel.Cell { Index = i, Name = Convert.ToString(e.Row.ItemArray[i]) });
                    }

                    sheet.Rows.Add(bexcelRow);
                    Debug.WriteLine($"New row added to {((dynamic)sender).TableName} table");
                }
                else
                {
                    var rowIndex = sheet.Dt.Rows.IndexOf(e.Row);

                    for (int i = 0; i < e.Row.ItemArray.Length; i++)
                    {
                        if (sheet.Rows[rowIndex].Cells.Find(x => x.Index == i) != null)
                        {
                            sheet.Rows[rowIndex].Cells[i].Name = Convert.ToString(e.Row.ItemArray[i]);
                        }
                        else
                        {
                            sheet.Rows[rowIndex].Cells.Add(new Definitions.Bexcel.Cell
                            {
                                Index = i,
                                Name = Convert.ToString(e.Row.ItemArray[i])
                            });
                        }
                    }

                    Debug.WriteLine($"Row updated on {((dynamic)sender).TableName} table");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error happened when updating the row.\n\n" + ex);
            }
        }

        private void Dt_RowDeleting(object sender, DataRowChangeEventArgs e)
        {
            try
            {
                var sheet = Bexcel.Sheets.First(x => x.Name == ((dynamic)sender).TableName);
                var rowIndex = sheet.Dt.Rows.IndexOf(e.Row);

                sheet.Rows.RemoveAt(rowIndex);

                Debug.WriteLine($"Row deleted on {((dynamic)sender).TableName} table");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error happened when deleting the row.\n\n" + ex);
            }
        }

        private void OpenBexcelFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".bexcel",
                Filter = "Bexcel file (*.bexcel)|*.bexcel"
            };

            if (dlg.ShowDialog() == true)
            {
                OpenBexcelFile(dlg.FileName);
            }
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateItemSource(SearchSheets.Text);
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Bexcel == null || Sheet.ItemsSource == null)
                    return;

                DataView dataViewSheet = (DataView)Sheet.ItemsSource;
                var sheet = Bexcel.Sheets.First(x => x.Name == dataViewSheet.Table.TableName);

                foreach (DataRow row in sheet.Dt.Rows)
                {
                    var rowIndex = sheet.Dt.Rows.IndexOf(row);

                    if (sheet.Rows.ElementAtOrDefault(rowIndex) == null)
                    {
                        var bexcelRow = new Definitions.Bexcel.Row
                        {
                            Index1 = -1,
                            Index2 = 1
                        };

                        for (int i = 0; i < row.ItemArray.Length; i++)
                        {
                            bexcelRow.Cells.Add(new Definitions.Bexcel.Cell { Name = (string)row.ItemArray[i] });
                        }

                        sheet.Rows.Add(bexcelRow);
                    }
                    else
                    {
                        for (int i = 0; i < row.ItemArray.Length; i++)
                        {
                            sheet.Rows[rowIndex].Cells[i].Name = (string)row.ItemArray[i];
                        }
                    }
                }

                MessageBox.Show("Data saved to memory");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error happened when updating the row.\n\n" + ex);
            }
        }
    }
}