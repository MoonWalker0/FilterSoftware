using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace TelevendFilter
{
    /// <summary>
    /// Logika interakcji dla klasy MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        string globalDatabasePath = null;
         
        public MainWindow()
        {
            InitializeComponent();
            base.Closing += this.MainWindow_Closing;
            ((INotifyCollectionChanged)MainList.Items).CollectionChanged += MainList_CollectionChanged;
            try
            {
                this.Title = "TeleVend Audit Filter - " + File.ReadAllText(@"Config.cfg");
            }
            catch(System.IO.FileNotFoundException)
            {
                this.Title = "TeleVend Audit Filter";
            }
        } 

        void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (globalDatabasePath != null)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                File.Delete(globalDatabasePath);
            }
        }

        private async void SelectFileButton(object sender, System.Windows.RoutedEventArgs e)
        {
            CommonOpenFileDialog dialogFileSelect = new CommonOpenFileDialog();
            dialogFileSelect.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialogFileSelect.Filters.Add(new CommonFileDialogFilter("Arkusz programu Microsoft Excel (*.xlsx)", ".xlsx"));
            if (dialogFileSelect.ShowDialog() == CommonFileDialogResult.Ok)
            {
                //Remove previous DB (if exists)
                if (globalDatabasePath != null)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    File.Delete(globalDatabasePath);
                }

                globalDatabasePath = dialogFileSelect.FileName.Remove(dialogFileSelect.FileName.Length - 4, 4) + "sqlite";
                
                //Get size and verify file
                int rowNumber = DatabaseHandling.GetRowNumber(globalDatabasePath);

                ////////////////////////////////////////////////////////////////////////
                //                              DEMO                                  //
                ////////////////////////////////////////////////////////////////////////

                //if (rowNumber > 50)
                //{
                //    await this.ShowMessageAsync("DEMO", "Wersja demo obsługuje do 50 wierszy pliku wsadowego.");
                //    globalDatabasePath = null;
                //    return;
                //}

                ////////////////////////////////////////////////////////////////////////
                //                              DEMO                                  //
                ////////////////////////////////////////////////////////////////////////


                if (rowNumber == -1)
                {
                    await this.ShowMessageAsync("Błąd pliku", "Załadowany plik jest niepoprawny");
                    globalDatabasePath = null;
                    return;
                }
                else if(rowNumber == -2)
                {
                    await this.ShowMessageAsync("Plik w użyciu", "Wybrany plik jest otwarty. Zamknij go i spróbuj ponownie.");
                    globalDatabasePath = null;
                    return;
                }

                //Create DB
                ProgressDialogController controller = await this.ShowProgressAsync("Ładowanie", "Dane są odczytywane");
                controller.Maximum = (double)rowNumber;
                controller.Minimum = 0;

                IProgress<double> progress = new Progress<double>(value => controller.SetProgress(value));
                progress.Report(0);

                await Task.Run(() =>
                {
                    DatabaseHandling.CreateTempDB(globalDatabasePath);
                    DatabaseHandling.FillDB(globalDatabasePath, progress);
                });
                await controller.CloseAsync(); 

                //Fill list 
                MainList.ItemsSource = DatabaseHandling.LoadAllByDate(globalDatabasePath);
                
                //Invalid file was loaded
                if(MainList.Items.Count == 0)
                {
                    await this.ShowMessageAsync("Błąd pliku", "Załadowany plik jest pusty lub niepoprawny.");
                    globalDatabasePath = null; 
                }
            }
            else
            {
                return;
            }
        }

        private void AddWorkerIDButton(object sender, System.Windows.RoutedEventArgs e)
        {
            if (globalDatabasePath != null)
            {
                CommonOpenFileDialog dialogFileSelect = new CommonOpenFileDialog();
                dialogFileSelect.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                dialogFileSelect.Filters.Add(new CommonFileDialogFilter("Arkusz programu Microsoft Excel (*.xlsx)", ".xlsx"));
                if (dialogFileSelect.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    //Update DB
                    if(DatabaseHandling.AddWorkerID(globalDatabasePath, dialogFileSelect.FileName) == 0)
                    {
                        this.ShowMessageAsync("Błąd pliku", "Wybrany plik jest niepoprawny, pusty lub otwarty");
                        return;
                    }

                    //Fill list 
                    MainList.ItemsSource = DatabaseHandling.LoadAllByDate(globalDatabasePath);
                    this.ShowMessageAsync("Gotowe", "Załadowano kodowanie.");
                }
            }
            else
            {
                this.ShowMessageAsync("Brak danych", "Dodaj najpierw plik");
            }
        }

        private void ExportDataButton(object sender, System.Windows.RoutedEventArgs e)
        {
            if (globalDatabasePath == null)
            {
                this.ShowMessageAsync("Brak danych", "Dodaj najpierw plik");
            }
            else if(MainList.Items.Count == 0)
            {
                this.ShowMessageAsync("Brak danych", "Lista jest pusta");
            }
            else
            {
                CommonSaveFileDialog dialogFileSelect = new CommonSaveFileDialog();
                dialogFileSelect.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                dialogFileSelect.Filters.Add(new CommonFileDialogFilter("Adobe Acrobat Document (*.pdf)", ".pdf"));
                dialogFileSelect.Filters.Add(new CommonFileDialogFilter("Arkusz programu Microsoft Excel (*.xlsx)", ".xlsx"));
                if (dialogFileSelect.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    if (dialogFileSelect.SelectedFileTypeIndex == 2)
                    {
                        ExcelExport.SaveBigAudit(globalDatabasePath, dialogFileSelect.FileName, (MainList.ItemsSource as IEnumerable<ListDisplay.ListItem>), 
                                                DateFromPicker.SelectedDate, DateToPicker.SelectedDate, PurchaseSum.Text);
                    }
                    else
                    {
                        PdfExport.SaveBigAudit(globalDatabasePath, dialogFileSelect.FileName, (MainList.ItemsSource as IEnumerable<ListDisplay.ListItem>),
                                                DateFromPicker.SelectedDate, DateToPicker.SelectedDate, PurchaseSum.Text);
                    }
                    this.ShowMessageAsync("Gotowe", "Plik wyeksportowano.");
                }
            }
        }

        private void ExportIndividualButton(object sender, System.Windows.RoutedEventArgs e)
        {
            if (globalDatabasePath == null)
            {
                this.ShowMessageAsync("Brak danych", "Dodaj najpierw plik");
            }
            else if (MainList.Items.Count == 0)
            {
                this.ShowMessageAsync("Brak danych", "Lista jest pusta");
            }
            else
            {
                CommonSaveFileDialog dialogFileSelect = new CommonSaveFileDialog();
                dialogFileSelect.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                dialogFileSelect.Filters.Add(new CommonFileDialogFilter("Adobe Acrobat Document (*.pdf)", ".pdf"));
                dialogFileSelect.Filters.Add(new CommonFileDialogFilter("Arkusz programu Microsoft Excel (*.xlsx)", ".xlsx"));
                if (dialogFileSelect.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    if (dialogFileSelect.SelectedFileTypeIndex == 2)
                    {
                        ExcelExport.SaveSmallAudit(globalDatabasePath, dialogFileSelect.FileName, (MainList.ItemsSource as IEnumerable<ListDisplay.ListItem>),
                                            DateFromPicker.SelectedDate, DateToPicker.SelectedDate, PurchaseSum.Text);
                    }
                    else
                    {
                        PdfExport.SaveSmallAudit(globalDatabasePath, dialogFileSelect.FileName, (MainList.ItemsSource as IEnumerable<ListDisplay.ListItem>),
                                            DateFromPicker.SelectedDate, DateToPicker.SelectedDate, PurchaseSum.Text);
                    }
                    this.ShowMessageAsync("Gotowe", "Plik wyeksportowano.");
                }
            } 

        }

        private void DateLastMonthClick(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DateTime.Today.Month == 1)
            {
                DateFromPicker.SelectedDate = new DateTime(DateTime.Today.Year - 1,
                                                           12,
                                                           1);
                DateToPicker.SelectedDate = new DateTime(DateTime.Today.Year - 1,
                                                         12,
                                                         DateTime.DaysInMonth(DateTime.Today.Year - 1,
                                                                              12));
            }
            else
            {
                DateFromPicker.SelectedDate = new DateTime(DateTime.Today.Year,
                                                           DateTime.Today.Month - 1,
                                                           1);
                DateToPicker.SelectedDate = new DateTime(DateTime.Today.Year,
                                                         DateTime.Today.Month - 1,
                                                         DateTime.DaysInMonth(DateTime.Today.Year,
                                                                              DateTime.Today.Month - 1));

            }
        }

        private void DateThisMonthClick(object sender, System.Windows.RoutedEventArgs e)
        {
            DateFromPicker.SelectedDate = new DateTime(DateTime.Today.Year,
                                                       DateTime.Today.Month,
                                                       1);
            DateToPicker.SelectedDate = new DateTime(DateTime.Today.Year,
                                                     DateTime.Today.Month,
                                                     DateTime.DaysInMonth(DateTime.Today.Year,
                                                                          DateTime.Today.Month));
        }

        private void DateThisYearClick(object sender, System.Windows.RoutedEventArgs e)
        {
            DateFromPicker.SelectedDate = new DateTime(DateTime.Today.Year,
                                                       1,
                                                       1);
            DateToPicker.SelectedDate = new DateTime(DateTime.Today.Year,
                                                     12,
                                                     DateTime.DaysInMonth(DateTime.Today.Year,
                                                                          1));
        }

        private void MainListSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (MainList.SelectedItem != null)
            {
                CardID.Text = (MainList.SelectedItem as ListDisplay.ListItem?).Value.ItemID.ToString();
            }
        }

        private void PerformFilterButton(object sender, System.Windows.RoutedEventArgs e)
        {
            if (globalDatabasePath != null )
            {
                var tempList = DatabaseHandling.PerformFilter(DateFromPicker.SelectedDate, DateToPicker.SelectedDate,
                                                              CardID.Text, ServiceCard.Text, globalDatabasePath);

                if (tempList != null)
                {
                    MainList.ItemsSource = tempList;
                }
                else
                {
                    this.ShowMessageAsync("Błąd danych", "Niepoprawne dane filtra");
                }
            }
            else
            {
                this.ShowMessageAsync("Brak danych", "Dodaj najpierw plik");
            }
        }


        private void MainList_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        { 
            PurchaseSum.Text = (MainList.ItemsSource as IEnumerable<ListDisplay.ListItem>).
                                Sum(item => Convert.ToDecimal(item.ItemPurchase)).ToString();

        }

        private void ResetFilterButton(object sender, System.Windows.RoutedEventArgs e)
        {
            DateFromPicker.SelectedDate = null;
            DateToPicker.SelectedDate = null;
            CardID.Text = "Kod karty";
            ServiceCard.Text = "Kod karty";

            if(globalDatabasePath != null)
            {
                MainList.ItemsSource = DatabaseHandling.LoadAllByDate(globalDatabasePath);
            }

        }
    }
}
