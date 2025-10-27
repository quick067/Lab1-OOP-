using MyExcelMAUIApp.Services;
using System.Text.Json;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace MyExcelMAUIApp.Models
{
    public class Spreadsheet
    {
        private readonly Dictionary<string, Cell> cells = new();
        public Cell GetOrCreateCell(string address)
        {
            if (!cells.ContainsKey(address))
            {
                cells[address] = new Cell();
            }
            return cells[address];
        }
        private void RecalculateAllCells()
        {
            const int maxIterations = 10;
            for (int i = 0; i < maxIterations; i++)
            {
                bool wasChanged = false;
                // 1. Створюємо копію поточних значень ЯК КОНТЕКСТ для цієї ітерації
                var valuesContext = cells.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Value);

                // 2. Створюємо копію ключів для безпечної ітерації
                var addressesToCalculate = cells.Keys.ToList();

                foreach (string address in addressesToCalculate)
                {
                    if (!cells.TryGetValue(address, out Cell currentCell)) continue;

                    object? oldValue = currentCell.Value; // Використовуємо object?
                    object? newValue;

                    if (string.IsNullOrWhiteSpace(currentCell.Expression))
                    {
                        newValue = null;
                    }
                    // 3. Передаємо СЛОВНИК valuesContext як контекст
                    else if (currentCell.Expression.StartsWith("="))
                    {
                        newValue = Calculator.Evaluate(currentCell.Expression.Substring(1), valuesContext);
                    }
                    else
                    {
                        if (System.Numerics.BigInteger.TryParse(currentCell.Expression, out var num))
                        {
                            newValue = num;
                        }
                        else
                        {
                            newValue = currentCell.Expression;
                        }
                    }
                    if (!Equals(oldValue, newValue))
                    {
                        currentCell.Value = newValue;
                        // 4. Оновлюємо КОНТЕКСТ для наступних обчислень В ЦІЙ ІТЕРАЦІЇ
                        valuesContext[address] = newValue;
                        wasChanged = true;
                    }
                }
                if (!wasChanged)
                {
                    break;
                }
            }
        }

        private void UpdateDependencies(Cell cell)
        {
            var visitor = new DependenciesVisitor();
            cell.Dependencies = visitor.FindDependencies(cell.Expression);
        }
        private bool CheckForCoherentLooping(string modifiedCellAddress)
        {
            var recursionStack = new HashSet<string>();
            var visited = new HashSet<string>();

            bool DFS(string currentAddress)
            {
                recursionStack.Add(currentAddress);
                visited.Add(currentAddress);
                if(cells.TryGetValue(currentAddress, out var cell))
                {
                    foreach(string dependencyAdress in cell.Dependencies)
                    {
                        if (!visited.Contains(dependencyAdress))
                        {
                            if (DFS(dependencyAdress))
                            {
                                return true;
                            }
                        }
                        else if (recursionStack.Contains(dependencyAdress))
                        {
                            return true;
                        }
                    }
                }

                recursionStack.Remove(currentAddress);
                return false;
            }
            return DFS(modifiedCellAddress);
        }

        public void SetExpression(string address, string expression)
        {
            Cell cell = GetOrCreateCell(address);
            cell.Expression = expression;

            UpdateDependencies(cell);
            if (CheckForCoherentLooping(address))
            {
                cell.Value = "#ЦИКЛ!";
                return;
            }
            RecalculateAllCells();
        }

        public void SetCellValueDirectly(string address, object? value)
        {
            Cell cell = GetOrCreateCell(address);
            cell.Value = value;
        }

        public IReadOnlyDictionary<string, object> GetCellValues()
        {
            return cells.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Value);
        }

        public void SaveToFile(string filePath)
        {
            var expressionsToSave = cells.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Expression);

            var options = new JsonSerializerOptions { WriteIndented = true };
            string jsonString = JsonSerializer.Serialize(expressionsToSave, options);
            File.WriteAllText(filePath, jsonString);
        }

        public void LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }

            string jsonString = File.ReadAllText(filePath);
            var loadedExpressions = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonString);

            if (loadedExpressions != null)
            {
                cells.Clear();
                foreach (var pair in loadedExpressions)
                {
                    Cell newCell = new Cell { Expression = pair.Value };
                    cells[pair.Key] = newCell;

                    UpdateDependencies(newCell);
                }

                bool cycleFoundOnLoad = false;
                foreach (string address in cells.Keys)
                {
                    if (CheckForCoherentLooping(address))
                    {
                        cells[address].Value = "#ЦИКЛ!";
                        cycleFoundOnLoad = true;
                    }
                }

                if (!cycleFoundOnLoad)
                {
                    RecalculateAllCells();
                }
            }
        }

        public void OnRowDeleted(int deletedRowNumber)
        {
            var addressesToRemove = cells.Keys
                .Where(addr => ParseAddress(addr, out _, out int row) && row == deletedRowNumber)
                .ToList();

            foreach (var address in addressesToRemove)
            {
                cells.Remove(address);
            }
            RecalculateAllCells();
        }

        public void OnColumnDeleted(string deletedColumnName)
        {
            var addressesToRemove = cells.Keys
                .Where(addr => ParseAddress(addr, out string col, out _) && col == deletedColumnName)
                .ToList();

            foreach (var address in addressesToRemove)
            {
                cells.Remove(address);
            }
            RecalculateAllCells();
        }

        private bool ParseAddress(string address, out string columnName, out int rowNumber)
        {
            columnName = string.Concat(address.TakeWhile(char.IsLetter));
            string rowStr = string.Concat(address.SkipWhile(char.IsLetter));
            return int.TryParse(rowStr, out rowNumber) && !string.IsNullOrEmpty(columnName);
        }
    }
}
