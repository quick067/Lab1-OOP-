using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using System.Collections.Generic;
using System.Linq;
using System.Numerics; 

namespace MyExcelMAUIApp.Services
{
    public class ExpressionEvaluator : ExcelGrammarBaseVisitor<object>
    {
        private readonly Dictionary<string, object> _cellValues;

        public ExpressionEvaluator(Dictionary<string, object> cellValues)
        {
            _cellValues = cellValues ?? new Dictionary<string, object>();
        }

        public override object VisitParse(ExcelGrammarParser.ParseContext context)
        {
            return Visit(context.expression());
        }

        public override object VisitNumberExpr(ExcelGrammarParser.NumberExprContext context)
        {
            return BigInteger.Parse(context.GetText());
        }

        public override object VisitBooleanExpr(ExcelGrammarParser.BooleanExprContext context)
        {
            return bool.Parse(context.GetText().ToLower());
        }

        public override object VisitIdentifierExpr(ExcelGrammarParser.IdentifierExprContext context)
        {
            string cellAddress = context.GetText().ToUpper(); 

            if (_cellValues.TryGetValue(cellAddress, out object? cellValue))
            {
                if (cellValue is BigInteger numericValue)
                {
                    return numericValue;
                }
                if(cellValue is bool boolValue)
                {
                    return boolValue;
                }
                if (cellValue is string stringValue)
                {
                    if (stringValue.StartsWith("#"))
                    {
                        return stringValue;
                    }
                    if (bool.TryParse(stringValue, out bool parseBool))
                    {
                        return parseBool;
                    }
                }
                return new BigInteger(0);
            }
            else
            {
                return "#ПОС!"; 
            }
        }

        public override object VisitParenthesizedExpr(ExcelGrammarParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override object VisitMultiplicativeExpr(ExcelGrammarParser.MultiplicativeExprContext context)
        {
            object leftOperand = Visit(context.expression(0));
            object rightOperand = Visit(context.expression(1));

            if (leftOperand is BigInteger leftNumber && rightOperand is BigInteger rightNumber)
            {
                if (context.op.Type == ExcelGrammarLexer.MULTIPLY)
                {
                    return leftNumber * rightNumber;
                }
                if (context.op.Type == ExcelGrammarLexer.DIVIDE)
                {
                    if (rightNumber == 0)
                    {
                        return "#Div/0!"; 
                    }
                    return leftNumber / rightNumber;
                }
            }
            return "#ЗНАЧ!";
        }

        public override object VisitAdditiveExpr(ExcelGrammarParser.AdditiveExprContext context)
        {
            object leftOperand = Visit(context.expression(0));
            object rightOperand = Visit(context.expression(1));

            if (leftOperand is BigInteger leftNumber && rightOperand is BigInteger rightNumber)
            {
                if (context.op.Type == ExcelGrammarLexer.ADD) return leftNumber + rightNumber;
                if (context.op.Type == ExcelGrammarLexer.SUBTRACT) return leftNumber - rightNumber;
            }
            return "#ЗНАЧ!";
        }

        public override object VisitRelationalExpr(ExcelGrammarParser.RelationalExprContext context)
        {
            object leftOperand = Visit(context.expression(0));
            object rightOperand = Visit(context.expression(1));

            if (leftOperand is BigInteger leftNumber && rightOperand is BigInteger rightNumber)
            {
                return context.op.Text switch
                {
                    "=" => leftNumber == rightNumber,
                    "<" => leftNumber < rightNumber,
                    ">" => leftNumber > rightNumber,
                    _ => false
                };
            }
            return "#ЗНАЧ!";
        }

        public override object VisitFunctionExpr(ExcelGrammarParser.FunctionExprContext context)
        {
            string functionName = context.FUNCTION_NAME().GetText().ToLowerInvariant();
            List<object> arguments = context.expression().Select(Visit).ToList();

            switch (functionName)
            {
                case "not":
                    if(arguments.Count != 1)
                        return "#АРГ!";

                    if(arguments[0] is bool boolValue)
                    {
                        return !boolValue;
                    }
                    return "#ЗНАЧ!";

                case "inc":
                case "dec":
                case "mmax":
                case "mmin":

                    List<BigInteger> numericArguments = new List<BigInteger>();
                    foreach(var arg in arguments)
                    {
                        if(arg is BigInteger number)
                        {
                            numericArguments.Add(number);
                        }
                        else
                        {
                            return "#ЗНАЧ!";
                        }
                    }

            if(functionName == "inc")
                    return numericArguments.Count == 1 ? numericArguments[0] + 1 : (object)"#AРГ!";
            if (functionName == "dec")
                    return numericArguments.Count == 1 ? numericArguments[0] - 1 : (object)"#AРГ!";
            if (functionName == "mmax")
                    return numericArguments.Any() ? numericArguments.Max() : (object)"#AРГ!";
            if (functionName == "mmin")
                    return numericArguments.Any() ? numericArguments.Min() : (object)"#AРГ!";

            break;

                default:
                    return "#ІМ'Я?";
            }
            return "#ERROR!";
        }
    }
}