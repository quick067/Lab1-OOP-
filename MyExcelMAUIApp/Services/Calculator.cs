using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using System.Collections.Generic;
using MyExcelMAUIApp.Models;

namespace MyExcelMAUIApp.Services
{
    public static class Calculator
    {
        public static object Evaluate(string expression, Dictionary<string, object?> cellValuesContext)
        {
            try
            {
                var charStream = new AntlrInputStream(expression);
                var lexer = new ExcelGrammarLexer(charStream);

                lexer.RemoveErrorListeners();
                lexer.AddErrorListener(ExcelGrammarErrorListener.Instance);

                var tokens = new CommonTokenStream(lexer);
                var parser = new ExcelGrammarParser(tokens);

                parser.RemoveErrorListeners();
                parser.AddErrorListener(ExcelGrammarErrorListener.Instance);

                IParseTree syntaxTree = parser.parse();

                var visitor = new ExpressionEvaluator(cellValuesContext);

                return visitor.Visit(syntaxTree);
            }
            catch (System.ArgumentException ex)
            {
                return $"#ERROR: {ex.Message}";
            }
            catch
            {
                return "#ERROR!";
            }
        }
    }
}