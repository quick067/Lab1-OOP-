using Microsoft.Maui.Controls;
using MyExcelMAUIApp.Models;
using System.IO;
using System.Linq; 

namespace MyExcelMAUIApp
{
    public partial class MainPage : ContentPage
    {
        const int CountColumn = 20;
        const int CountRow = 50;

        private readonly Spreadsheet spreadsheet = new();

        private readonly Dictionary<string, Entry> uiCells = new();
        private string currentSelectedCellAddress = null;

        public MainPage()
        {
            InitializeComponent();
            CreateGrid();
        }

        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }

        private void AddColumnsAndColumnLabels()
        {
            for (int col = 0; col < CountColumn + 1; col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                if (col > 0)
                {
                    var label = new Label
                    {
                        Text = GetColumnName(col),
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(label, 0);
                    Grid.SetColumn(label, col);
                    grid.Children.Add(label);
                }
            }
        }

        private void AddRowsAndCellEntries()
        {
            for (int row = 0; row < CountRow; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition());
                var label = new Label
                {
                    Text = (row + 1).ToString(),
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                Grid.SetRow(label, row + 1);
                Grid.SetColumn(label, 0);
                grid.Children.Add(label);
                for (int col = 0; col < CountColumn; col++)
                {
                    var entry = new Entry
                    {
                        Text = "",
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    entry.Unfocused += Entry_Unfocused;
                    entry.Focused += Entry_Focused;
                    entry.Completed += Entry_Completed;
                    Grid.SetRow(entry, row + 1);
                    Grid.SetColumn(entry, col + 1);
                    grid.Children.Add(entry);
                    string cellAddress = GetColumnName(col + 1) + (row + 1);
                    uiCells[cellAddress] = entry;
                }
            }
        }

        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private void HandleUpdateOfCell(object? sender)
        {
            if (sender is not Entry entry) return;

            int rowIndex = Grid.GetRow(entry) - 1;
            int colIndex = Grid.GetColumn(entry) - 1;
            string newExpression = entry.Text;
            string cellAddress = GetColumnName(colIndex + 1) + (rowIndex + 1);

            spreadsheet.SetExpression(cellAddress, newExpression);
            updateGridDisplay();
        }

        private void Entry_Unfocused(object? sender, FocusEventArgs e)
        {
            HandleUpdateOfCell(sender);
        }

        private void Entry_Focused(object? sender, FocusEventArgs e)
        {
            if (sender is not Entry entry)
            {
                return;
            }
            int rowIndex = Grid.GetRow(entry) - 1;
            int colIndex = Grid.GetColumn(entry) - 1;
            string cellAddress = GetColumnName(colIndex + 1) + (rowIndex + 1);

            currentSelectedCellAddress = cellAddress;
            var cellModel = spreadsheet.GetOrCreateCell(cellAddress);

            textInput.Text = cellModel.Expression;

            if (cellModel.Value is string valueString && valueString.StartsWith("#"))
            {
                entry.Text = valueString;
            }
            else
            {
                entry.Text = cellModel.Expression;
            }
        }

        private void Entry_Completed(object? sender, EventArgs e)
        {
            HandleUpdateOfCell(sender);
            grid.Focus();
        }

        private void FormulaBar_Completed(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentSelectedCellAddress))
            {
                return;
            }
            string newExpression = textInput.Text;
            spreadsheet.SetExpression(currentSelectedCellAddress, newExpression);
            updateGridDisplay();
            grid.Focus();
        }
        private void updateGridDisplay()
        {
            var values = spreadsheet.GetCellValues();
            foreach (var uiCellPair in uiCells)
            {
                string address = uiCellPair.Key;
                Entry entryField = uiCellPair.Value;
                if (entryField.IsFocused) continue;
                if (values.TryGetValue(address, out object value) && value != null)
                {
                    entryField.Text = value.ToString();
                }
                else
                {
                    entryField.Text = "";
                }
            }
        }
        private async void SaveButton_Clicked(object? sender, EventArgs e)
        {
            try
            {
                DateTime now = DateTime.Now;
                string time = now.ToString("yyyy-MM-dd_HH-mm-ss");
                string fileName = $"MyExcelSheet_{time}.json";
                string filePath = Path.Combine(FileSystem.AppDataDirectory, fileName);
                spreadsheet.SaveToFile(filePath);
                await DisplayAlert("Успіх", $"Файл збережено як: {fileName}", "OK");
            }
            catch (Exception ex) { await DisplayAlert("Помилка", ex.Message, "OK"); }
        }

        private async void ReadButton_Clicked(object? sender, EventArgs e)
        {
            try
            {
                string appDataDir = FileSystem.AppDataDirectory;

                string[] savedFilesPaths = Directory.GetFiles(appDataDir, "MyExcelSheet_*.json");

                if (savedFilesPaths.Length == 0)
                {
                    await DisplayAlert("Не знайдено", "Збережених файлів не знайдено.", "OK");
                    return;
                }

                string[] fileNames = savedFilesPaths.Select(Path.GetFileName).ToArray();

                string selectedFileName = await DisplayActionSheet(
                    "Оберіть файл для відкриття", 
                    "Скасувати",                  
                    null,                         
                    fileNames);                 

                if (string.IsNullOrEmpty(selectedFileName) || selectedFileName == "Скасувати")
                {
                    return;
                }

                string selectedFilePath = Path.Combine(appDataDir, selectedFileName);

                spreadsheet.LoadFromFile(selectedFilePath);
                updateGridDisplay();

                await DisplayAlert("Успіх", $"Таблицю '{selectedFileName}' завантажено.", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка", ex.Message, "OK");
            }
        }

        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count <= 1) return;
            int rowToDelete = grid.RowDefinitions.Count;
            spreadsheet.OnRowDeleted(rowToDelete);
            var elementsToRemove = grid.Children.Where(c => Grid.GetRow((BindableObject)c) == rowToDelete - 1).ToList();
            foreach (var element in elementsToRemove) { grid.Children.Remove(element); }
            grid.RowDefinitions.RemoveAt(rowToDelete - 1);
            updateGridDisplay();
        }

        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count <= 1) return;
            int columnToDelete = grid.ColumnDefinitions.Count - 1;
            spreadsheet.OnColumnDeleted(GetColumnName(columnToDelete));
            var elementsToRemove = grid.Children.Where(c => Grid.GetColumn((BindableObject)c) == columnToDelete).ToList();
            foreach (var element in elementsToRemove) { grid.Children.Remove(element); }
            grid.ColumnDefinitions.RemoveAt(columnToDelete);
            updateGridDisplay();
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            int newRow = grid.RowDefinitions.Count;
            grid.RowDefinitions.Add(new RowDefinition());
            var label = new Label
            {
                Text = newRow.ToString(),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, newRow);
            Grid.SetColumn(label, 0);
            grid.Children.Add(label);
            for (int col = 0; col < CountColumn; col++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, newRow);
                Grid.SetColumn(entry, col + 1);
                grid.Children.Add(entry);
            }
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            int newColumn = grid.ColumnDefinitions.Count;
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            var label = new Label
            {
                Text = GetColumnName(newColumn),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, newColumn);
            grid.Children.Add(label);
            for (int row = 0; row < CountRow; row++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, row + 1);
                Grid.SetColumn(entry, newColumn);
                grid.Children.Add(entry);
            }
        }

        private void CalculateButton_Clicked(object sender, EventArgs e)
        {
            updateGridDisplay();
            DisplayAlert("Готово", "Відображення таблиці оновлено.", "OK");
        }
        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Лабораторна робота 1. Студента Миколи Шевченка", "OK");
        }

        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                System.Environment.Exit(0);
            }
        }
    }
}