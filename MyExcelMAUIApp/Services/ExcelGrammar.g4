grammar ExcelGrammar;


parse : expression EOF;

expression
    : expression op=('=' | '<' | '>') expression   #RelationalExpr 
    | expression op=('+' | '-') expression         #AdditiveExpr  
    | expression op=('*' | '/') expression         #MultiplicativeExpr
    | '-' expression                               #UnaryMinusExpr 
    | FUNCTION_NAME LPAREN expression (',' expression)* RPAREN #FunctionExpr  
    | IDENTIFIER                                   #IdentifierExpr 
    | BOOLEAN                                      #BooleanExpr
    | NUMBER                                       #NumberExpr     
    | LPAREN expression RPAREN                     #ParenthesizedExpr 
    ;



FUNCTION_NAME: ('i' 'n' 'c' | 'd' 'e' 'c' | 'm' 'm' 'a' 'x' | 'm' 'm' 'i' 'n' | 'n' 'o' 't');
BOOLEAN: ('T' 'R' 'U' 'E' | 'F' 'A' 'L' 'S' 'E' | 't' 'r' 'u' 'e' | 'f' 'a' 'l' 's' 'e');

NUMBER: [0-9]+;

IDENTIFIER: [A-Z]+[1-9][0-9]*;

EQUAL: '=';
LESS: '<';
GREATER: '>';
ADD: '+';
SUBTRACT: '-';
MULTIPLY: '*';
DIVIDE: '/';

LPAREN: '(';
RPAREN: ')';
COMMA: ',';

WS: [ \t\r\n]+ -> skip;