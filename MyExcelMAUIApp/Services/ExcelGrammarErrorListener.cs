using Antlr4.Runtime;
using Antlr4.Runtime.Misc;

namespace MyExcelMAUIApp.Services
{
    public class ExcelGrammarErrorListener : BaseErrorListener, IAntlrErrorListener<int>
    {
        public static readonly ExcelGrammarErrorListener Instance = new ExcelGrammarErrorListener();

        public override void SyntaxError([NotNull] IRecognizer recognizer, [Nullable] IToken offendingSymbol, int line, int charPositionInLine, [NotNull] string msg, [Nullable] RecognitionException e)
        {
            throw new System.ArgumentException($"Error: {msg}");
        }

        public void SyntaxError([NotNull] IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, [NotNull] string msg, [Nullable] RecognitionException e)
        {
            throw new System.ArgumentException($"Uknown: {msg}");
        }
    }
}