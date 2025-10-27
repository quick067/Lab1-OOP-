using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using System.Collections.Generic;
using System.Linq;

namespace MyExcelMAUIApp.Services
{
    class DependenciesVisitor : ExcelGrammarBaseVisitor<object>
    {
        private HashSet<string> dependencies = new HashSet<string>();
        public List<string> FindDependencies(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression) || !expression.StartsWith("="))
            {
                return new List<string>();
            }

            var charStream = new AntlrInputStream(expression.Substring(1));
            var lexer = new ExcelGrammarLexer(charStream);
            lexer.RemoveErrorListeners();
            var tokens = new CommonTokenStream(lexer);
            var parser = new ExcelGrammarParser(tokens);
            parser.RemoveErrorListeners();
            IParseTree tree = parser.parse();
            Visit(tree);
            return dependencies.ToList();
        }

        public override object VisitIdentifierExpr(ExcelGrammarParser.IdentifierExprContext context)
        {
            dependencies.Add(context.GetText().ToUpper());
            return null;
        }
        public override object VisitParse(ExcelGrammarParser.ParseContext context)
        {
            return Visit(context.expression());
        }
        public override object VisitParenthesizedExpr(ExcelGrammarParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }
        public override object VisitRelationalExpr(ExcelGrammarParser.RelationalExprContext context)
        {
            Visit(context.expression(0));
            Visit(context.expression(1));
            return null;
        }
        public override object VisitAdditiveExpr(ExcelGrammarParser.AdditiveExprContext context)
        {
            Visit(context.expression(0));
            Visit(context.expression(1));
            return null;
        }
        public override object VisitMultiplicativeExpr(ExcelGrammarParser.MultiplicativeExprContext context)
        {
            Visit(context.expression(0));
            Visit(context.expression(1));
            return null;
        }
        public override object VisitFunctionExpr(ExcelGrammarParser.FunctionExprContext context)
        {
            foreach (var expr in context.expression())
            {
                Visit(expr);
            }
            return null;
        }
    }
}
