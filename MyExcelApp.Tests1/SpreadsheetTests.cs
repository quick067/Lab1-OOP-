using Microsoft.VisualStudio.TestTools.UnitTesting;
using MyExcelMAUIApp.Models; 
using System.Numerics;       

namespace MyExcelApp.Tests
{
    [TestClass] 
    public class SpreadsheetTests
    {
        [TestMethod] 
        public void Calculate_SimpleDependency_ReturnsCorrectValue()
        {
            // Arrange (Підготовка)
            var sheet = new Spreadsheet();
            sheet.SetExpression("A1", "10");
            sheet.SetExpression("A2", "5");
            sheet.SetExpression("B1", "=A1+A2");

            // Act (Дія)
            var values = sheet.GetCellValues();
            values.TryGetValue("B1", out object? result); 

            Assert.IsNotNull(result, "Результат для B1 не повинен бути null.");
            Assert.IsInstanceOfType(result, typeof(BigInteger), "Результат має бути типу BigInteger.");
            Assert.AreEqual(new BigInteger(15), (BigInteger)result, "Значення B1 має бути 15.");
        }


        [TestMethod]
        public void Calculate_UpdateDependency_UpdatesDependentCell()
        {
            var sheet = new Spreadsheet();
            sheet.SetExpression("A1", "10");
            sheet.SetExpression("B1", "=A1*2"); 
            var initialValues = sheet.GetCellValues();
            initialValues.TryGetValue("B1", out object? initialResult);
            Assert.AreEqual(new BigInteger(20), (BigInteger?)initialResult, "Початкове значення B1 має бути 20.");

            sheet.SetExpression("A1", "100");
            var updatedValues = sheet.GetCellValues();
            updatedValues.TryGetValue("B1", out object? updatedResult); 

            Assert.IsNotNull(updatedResult, "Оновлений результат для B1 не повинен бути null.");
            Assert.AreEqual(new BigInteger(200), (BigInteger?)updatedResult, "Оновлене значення B1 має бути 200.");
        }


        [TestMethod]
        public void SetExpression_CircularDependency_SetsErrorValue()
        {
            var sheet = new Spreadsheet();
            sheet.SetExpression("X1", "=Y1"); 
            sheet.SetExpression("Y1", "=X1");

            var values = sheet.GetCellValues();
            values.TryGetValue("Y1", out object? resultY1); 

            values.TryGetValue("X1", out object? resultX1);

            Assert.IsNotNull(resultY1, "Результат для Y1 не повинен бути null.");
            Assert.AreEqual("#ЦИКЛ!", resultY1 as string, "Значення Y1 має бути '#ЦИКЛ!'.");
        }
    }
}