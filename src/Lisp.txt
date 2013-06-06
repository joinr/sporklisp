'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

Option Explicit

Private n As Long

'
'################ Lispy: Scheme Interpreter in Python
'
'## (c) Peter Norvig, 2010; See http://norvig.com/lispy.html
'
'################ Symbol, Env classes
'
'from __future__ import division
'
'symbol = str
'
'Class env(dict):
'    "An environment: a dict of {'var':val} pairs, with an outer Env."
'    def __init__(self, parms=(), args=(), outer=None):
'        self.update (zip(parms, args))
'        self.outer = outer
'    def Find(self, Var):
'        "Find the innermost Env where var appears."
'        return self if var in self else self.outer.find(var)
'
Const outer As String = "Outer"
Const lispPI As Double = 3.14159265358979
Const lispE As Double = 2.71828182845905

Public Enum expressionTypes
    selfEvaluatingQ = 0
    quotedQ = 1
    variableQ = 2
    assignmentQ = 3
    definitionQ = 4
    ifQ = 5
    lambdaQ = 6
    beginQ = 7
    condQ = 8
    applicationQ = 9
    unknown = -1
End Enum
    
Public globalenv As Dictionary
Public primitiveFuncs As Dictionary

Public Function getGlobalLispEnv() As Dictionary
If globalenv Is Nothing Then
    Set globalenv = installPrimitives(makeEnvironment)
    Set globalenv = installLibraries(globalenv)
End If
Set getGlobalLispEnv = globalenv
End Function

Public Function makePrimitive(id As Long, minArgs As Long, Optional maxArgs As Long) As LispPrimitive
Set makePrimitive = New LispPrimitive

If maxArgs = 0 Then maxArgs = minArgs
With makePrimitive
    .idNumber = id
    .minArgs = minArgs
    .maxArgs = maxArgs
End With

End Function
'define something in the - likely global - environment
Public Sub addLib(expr As String, Optional env As Dictionary)
If env Is Nothing Then Set env = getGlobalLispEnv
Call eval(expr, env)
End Sub
'equivalent to Norvig's python class.
Public Function makeEnvironment(Optional params As Collection, Optional args As Collection, _
                                    Optional outerenv As Dictionary) As Dictionary
Set makeEnvironment = DictionaryLib.zipMap(params, args)
If exists(outerenv) Then makeEnvironment.add outer, outerenv

End Function

Public Function withLisp(Optional params As Collection, Optional args As Collection) As Dictionary
Set withLisp = makeEnvironment(params, args, Lisp.getGlobalLispEnv)
End Function

'similar to norvig's Find, looks for inner most environment to allow lexical scoping.
Public Function findOuter(env As Dictionary, ByRef var As String) As Dictionary
Dim e As Dictionary

Set e = env
Do
    If e.exists(var) Then
        Set findOuter = e
        Exit Do
    ElseIf e.exists(outer) Then
        Set e = e(outer)
    Else
        Exit Do
    End If
Loop

End Function

'need to figure how to wrap this...

'def add_globals(env):
'    "Add some Scheme standard procedures to an environment."
'    import math, operator as op
'    env.update(vars(math)) # sin, sqrt, ...
'    env.update(
'     {'+':op.add, '-':op.sub, '*':op.mul, '/':op.div, 'not':op.not_,
'      '>':op.gt, '<':op.lt, '>=':op.ge, '<=':op.le, '=':op.eq,
'      'equal?':op.eq, 'eq?':op.is_, 'length':len, 'cons':lambda x,y:[x]+y,
'      'car':lambda x:x[0],'cdr':lambda x:x[1:], 'append':op.add,
'      'list':lambda *x:list(x), 'list?': lambda x:isa(x,list),
'      'null?':lambda x:x==[], 'symbol?':lambda x: isa(x, Symbol)})
'    return env
'
'global_env = add_globals(env())
'
'isa = isinstance


Public Function isa(var, ByRef nm As String) As Boolean
isa = TypeName(var) = nm
End Function

'from JScheme
' //////////////// Evaluation
'
'  /** Evaluate an object, x, in an environment. **/
'  public Object eval(Object x, Environment env) {
'    // The purpose of the while loop is to allow tail recursion.
'    // The idea is that in a tail recursive position, we do "x = ..."
'    // and loop, rather than doing "return eval(...)".
'    while (true) {
'      if (x instanceof String) {         // VARIABLE
'    return env.lookup((String)x);
'      } else if (!(x instanceof Pair)) { // CONSTANT
'    return x;
'      } else {
'    Object fn = first(x);
'    Object args = rest(x);
'    if (fn == "quote") {             // QUOTE
'      return first(args);
'    } else if (fn == "begin") {      // BEGIN
'      for (; rest(args) != null; args = rest(args)) {
'        eval(first(args), env);
'      }
'      x = first(args);
'    } else if (fn == "define") {     // DEFINE
'      if (first(args) instanceof Pair)
'        return env.define(first(first(args)),
'         eval(cons("lambda", cons(rest(first(args)), rest(args))), env));
'      else return env.define(first(args), eval(second(args), env));
'    } else if (fn == "set!") {       // SET!
'      return env.set(first(args), eval(second(args), env));
'    } else if (fn == "if") {         // IF
'      x = (truth(eval(first(args), env))) ? second(args) : third(args);
'    } else if (fn == "cond") {       // COND
'      x = reduceCond(args, env);
'    } else if (fn == "lambda") {     // LAMBDA
'      return new Closure(first(args), rest(args), env);
'    } else if (fn == "macro") {      // MACRO
'      return new Macro(first(args), rest(args), env);
'    } else {                         // PROCEDURE CALL:
'      fn = eval(fn, env);
'      if (fn instanceof Macro) {          // (MACRO CALL)
'        x = ((Macro)fn).expand(this, (Pair)x, args);
'      } else if (fn instanceof Closure) { // (CLOSURE CALL)
'        Closure f = (Closure)fn;
'        x = f.body;
'        env = new Environment(f.parms, evalList(args, env), f.env);
'      } else {                            // (OTHER PROCEDURE CALL)
'        return Procedure.proc(fn).apply(this, evalList(args, env));
'      }
'    }
'      }
'    }
'  }

'################ eval
'
'def eval(x, env = global_env):
'    "Evaluate an expression in an environment."
'    if isa(x, Symbol):             # variable reference
'        return env.find(x)[x]
'    elif not isa(x, list):         # constant literal
'        return x
'    elif x[0] == 'quote':          # (quote exp)
'        (_, exp) = x
'        return exp
'    elif x[0] == 'if':             # (if test conseq alt)
'        (_, test, conseq, alt) = x
'        return eval((conseq if eval(test, env) else alt), env)
'    elif x[0] == 'set!':           # (set! var exp)
'        (_, var, exp) = x
'        env.find(var)[var] = eval(exp, env)
'    elif x[0] == 'define':         # (define var exp)
'        (_, var, exp) = x
'        env [Var] = eval(Exp, env)
'    elif x[0] == 'lambda':         # (lambda (var*) exp)
'        (_, vars, exp) = x
'        return lambda *args: eval(exp, Env(vars, args, env))
'    elif x[0] == 'begin':          # (begin exp*)
'        for exp in x[1:]:
'            val = eval(Exp, env)
'        return val
'    else:                          # (proc exp*)
'        exps = [eval(exp, env) for exp in x]
'        proc = exps.pop(0)
'        return proc(*exps)
Function truth(inval As Variant) As Boolean
Select Case vartype(inval)
    Case VbVarType.vbBoolean, VbVarType.vbByte, VbVarType.vbInteger
        truth = CBool(inval)
    Case VbVarType.vbString
        If Mid(CStr(inval), 1, 1) = ":" Then
            truth = True
        Else
            truth = CBool(inval)
        End If
    Case Else
        truth = exists(inval)
End Select

End Function

Function isRange(inval As Variant) As Boolean
isRange = TypeName(inval) = "Range"
End Function

Function asRange(inval As Variant) As Range
Set asRange = inval
End Function

Function rangeValue(inval As Range) As Variant
Dim tmp() As Variant

If inval.Cells.count = 1 Then
    bind rangeValue, inval.value
Else
    tmp = inval
    bind rangeValue, CollectionLib.varrayToColl(tmp)
End If

End Function
Function checkRange(inval As Variant) As Variant

If isRange(inval) Then
    bind checkRange, rangeValue(asRange(inval))
Else
    bind checkRange, inval
End If

End Function
Sub resetLisp()
Set globalenv = Nothing
Set Lisp.primitiveFuncs = Nothing
End Sub



'Side-effecting way to bind vars
Sub bind(invar As Variant, ByRef inval As Variant)

If IsObject(inval) Then
    Set invar = inval
Else
    invar = inval
End If

End Sub
Sub clearbind(ByRef invar As Variant)
If IsObject(invar) Then
    Set invar = Nothing
Else
    invar = Empty
End If
End Sub
Function lisperror(expr As Variant, msg As String) As Collection
Set lisperror = list("error", msg)
End Function
Function unboundError(v As String) As String
unboundError = "No value defined for var " & v
End Function

Function eval(expr As Variant, Optional env As Dictionary) As Variant
Dim var As String
Dim vars As Collection
Dim outerd As Dictionary
Dim x As Variant

If env Is Nothing Then Set env = getGlobalLispEnv()

If isa(expr, "String") Then 'string literal, indicates a variable ref
    var = CStr(expr)
    If isLiteral(var) Then
        If Mid(var, 1, 1) = ":" Then
            bind eval, Mid(var, 1, Len(var))  'drops the outer set of " from the string
        Else
            bind eval, Mid(var, 2, Len(var) - 2) 'drops the outer set of " from the string
        End If
    ElseIf isBool(CStr(expr)) Then
        bind eval, CBool(expr)
    Else
        Set outerd = findOuter(env, var)
        If outerd Is Nothing Then
            'failed to find a var
            'signal an error condition in evaluation
            'bind eval, lisperror(expr, unboundError(var))
            Err.Raise 101, , unboundError(var)
        Else
            bind eval, outerd(var)
            Set outerd = Nothing
        End If
    End If
ElseIf Not isa(expr, "Collection") Then 'not a list, must be a constant/primitive
    eval = expr
Else
    Dim explist As Collection
    Set explist = expr
    If Not isa(explist(1), "Collection") Then
        Select Case explist(1)
            Case "quote" 'quoted expression
                If isLiteral(CStr(explist(2))) Then
                    bind eval, expr
                Else
                    bind eval, explist(2)
                End If
            Case "if"
                If truth(eval(explist(2), env)) Then
                    bind eval, eval(explist(3), env)
                Else
                    bind eval, eval(explist(4), env)
                End If
            Case "set!"
                var = CStr(explist(2))
                '(_, var, exp) = x
                'env.find(var)[var] = eval(exp, env)
                Dim val As Variant
                With findOuter(env, var)
                     bind val, eval(explist(3), env)
                    .Remove var
                    .add var, val
                End With
                clearbind val
            Case "let", "let*"
                bind eval, eval(letStar(explist), env)
            Case "define", "def"
                If Not isa(explist(2), "Collection") Then
                    var = CStr(explist(2))
                    env.add var, eval(explist(3), env)
    
                    eval = var
                Else 'implicit function definition
                    Set vars = explist(2) 'expand the (name par1 par2)
                    var = vars(1)
                    If env.exists(var) Then env.Remove (var) 'this might be buggy.
                    env.add var, makeClosure(restList(explist(2)), explist(3), env)
                    eval = var
                End If
            '    elif x[0] == 'define':         # (define var exp)
            '        (_, var, exp) = x
            '        env [Var] = eval(Exp, env)
            Case "defn" '(defn name [args] body)
                var = explist(2)
                Set vars = explist(3) 'expand the
                If env.exists(var) Then env.Remove (var) 'this might be buggy.
                env.add var, makeClosure(vars, explist(4), env)
                eval = var
            Case "lambda", "fn" 'this doesn't work, since we don't have lambdas...using jscheme example
            '    } else if (fn == "lambda") {     // LAMBDA
            '      return new Closure(first(args), rest(args), env);
                Set eval = makeClosure(explist(2), explist(3), env)
            Case "begin", "do", "progn"
            '    elif x[0] == 'begin':          # (begin exp*)
            '        for exp in x[1:]:
            '            val = eval(Exp, env)
            '        return val
                Dim v As Variant
                Dim j As Long
                
                For j = 2 To explist.count
                    bind v, eval(explist(j), env)
                Next j
                
                bind eval, v
                clearbind v
            Case "error"
                Err.Raise 101, , explist(2)
            Case Else
                bind eval, applyfn(explist, env)
        End Select
    Else
    '    } else {                         // PROCEDURE CALL:
    '      fn = eval(fn, env);
''               Dim fn As LispClosure
''               Dim i As Long
        bind eval, applyfn(explist, env)
       
    '      if (fn instanceof Macro) {          // (MACRO CALL)
    '        x = ((Macro)fn).expand(this, (Pair)x, args);
    '      } else if (fn instanceof Closure) { // (CLOSURE CALL)
    '        Closure f = (Closure)fn;
    '        x = f.body;
    '        env = new Environment(f.parms, evalList(args, env), f.env);
    '      } else {                            // (OTHER PROCEDURE CALL)
    '        return Procedure.proc(fn).apply(this, evalList(args, env));
    '      }
    '    }
    '    else:                          # (proc exp*)
    '        exps = [eval(exp, env) for exp in x]
    '        proc = exps.pop(0)
    '        return proc(*exps)
        
    End If
End If
    
End Function
Function isBool(ByRef s As String) As Boolean
Select Case UCase(s)
    Case "TRUE", "FALSE"
        isBool = True
    Case Else
        isBool = False
End Select
End Function
Function isLiteral(ByRef s As String) As Boolean
Select Case firstchar(s)
    Case ":"
        isLiteral = True
    Case Chr(34)
        isLiteral = True
    Case Else
        isLiteral = False
End Select
End Function
Function applyfn(explist As Collection, env As Dictionary) As Variant

Dim fn As IFn
Dim lc As LispClosure
Dim xs As Collection
Set xs = New Collection

Dim i As Long
Set fn = eval(explist(1), env)
For i = 2 To explist.count
    xs.add eval(explist(i), env)
Next i

bind applyfn, fn.apply(xs) 'res(1)
Set xs = Nothing
End Function
'destructures the explist, assumed to be the form
'(let
'   ((b1 v1) (b2 v2) (b3 v3))
'   body)
Function getbinds(explist As Collection) As Collection
Dim body As Variant
Dim binds As Collection
Set binds = explist(2)
End Function
Function getVarsVals(binds As Collection) As Collection
Dim itm As Collection
Dim vars As Collection
Dim vals As Collection

Set vars = New Collection
Set vals = New Collection

For Each itm In binds
    vars.add itm(1)
    vals.add itm(2)
Next itm
     
Set getVarsVals = list(vars, vals)

Set vars = Nothing
Set vals = Nothing

End Function
'assumes input : (let ((a 1) (b 2)) (+ a b))
Function letStar(explist As Collection)
Dim varvals As Collection
Set varvals = getVarsVals(explist(2))
Set letStar = nestedlet(varvals(1), varvals(2), explist(3))
End Function
'syntactic transform
'(let ( ((a 1)
'        (b 2))
'  (+ a b))
'becomes
'((lambda (a)
'   (+ a b)) 1 2)
Function deflet(parameters As Collection, args As Collection, body As Variant) As Collection
Dim itm As Variant
If args.count <> parameters.count Then Err.Raise 101, , "let requires an even number of binding forms"

Set deflet = list(list("lambda", parameters, body))
For Each itm In args
    deflet.add itm
Next itm

End Function
Function nestedlet(parameters As Collection, args As Collection, body As Variant) As Collection
Dim itm As Variant
Dim i As Long
Dim outerlist As Collection
Dim nextlist As Collection
If args.count <> parameters.count Then Err.Raise 101, , "let requires an even number of binding forms"

Set nestedlet = New Collection

Set nestedlet = list(list("lambda", list(parameters(parameters.count)), body), args(args.count))
For i = args.count - 1 To 1 Step -1
    Set outerlist = list(list("lambda", list(parameters(i)), nestedlet), args(i))
    Set nestedlet = outerlist
Next i

Set outerlist = Nothing
End Function

Public Function evalList(args As Collection, env As Dictionary) As Variant

End Function
Public Function makeClosure(params As Collection, body As Variant, env As Dictionary) As LispClosure
Set makeClosure = New LispClosure
With makeClosure
    Set .params = params
    If IsObject(body) Then
        Set .body = body
    Else
        .body = body
    End If
        
    Set .env = env
End With

End Function
Public Function read(ByRef s As String) As Variant
bind read, readfrom(tokenize(s))
End Function
'define anonymous function
Public Function makefn(ByRef expr As String, Optional env As Dictionary) As LispClosure
Set makefn = eval(read(expr), env)
End Function

'################ parse, read, and user interaction
'
'def Read(s):
'    "Read a Scheme expression from a string."
'    return read_from(tokenize(s))
'
'parse = Read
'
'def tokenize(s):
'    "Convert a string into a list of tokens."
'    return s.replace('(',' ( ').replace(')',' ) ').split()
'
Function tokenize(ByRef s As String) As Variant

Dim tmp
Dim tok As String
Dim itm
tok = replace(s, "(", " ( ")
tok = replace(tok, ")", " ) ")
tok = replace(tok, "[", " [ ")
tok = replace(tok, "]", " ] ")
tok = replace(tok, "{", " { ")
tok = replace(tok, "}", " } ")
tok = replace(tok, ",", " , ")

tmp = Split(tok)
'If UBound(tmp, 1) = 0 Then
'    tokenize = tmp(0)
'Else
    Set tokenize = New Collection
    For Each itm In tmp
        tokenize.add itm
    Next itm
'End If

End Function

'def read_from(tokens):
'    "Read an expression from a sequence of tokens."
'    if len(tokens) == 0:
'        raise SyntaxError('unexpected EOF while reading')
'    token = tokens.pop(0)
'    if '(' == token:
'        l = []
'        while tokens[0] != ')':
'            l.append (read_from(tokens))
'        tokens.pop(0) # pop off ')'
'        return L
'    elif ')' == token:
'        raise SyntaxError('unexpected )')
'    Else:
'        return atom(token)

Function readfrom(tokens As Collection, Optional ByRef delimiter As String) As Variant
Dim token As String
Dim l As Collection

If tokens.count = 0 Then _
    Err.Raise 101, "EOF While reading lisp expression!"

While whitespace(tokens(1)) And tokens.count > 0
    tokens.Remove 1
Wend
If tokens.count = 0 Then _
    Err.Raise 101, "EOF While reading lisp expression!"

token = tokens(1)

If isSequenceStart(token) Then
    delimiter = token
    Set l = New Collection
    tokens.Remove 1
    token = tokens(1)
    If delimiter <> "(" And token <> "hash-set" Then _
        l.add getSeqType(delimiter) 'prepends a constructor [vec, hash-map, hash-set]
        'we exclude list from here, since lists are treated as quoted lists
    
    While Not isSequenceStop(token, delimiter)
        If isSequenceStart(token) Or specialToken(token) Then
            l.add readfrom(tokens)
        ElseIf Not (whitespace(token)) Then
            l.add readfrom(tokens)
        End If
        tokens.Remove 1
        If tokens.count = 0 Then
            Err.Raise 101, , " missing " & getClosingDelim(delimiter)
        Else
            token = tokens(1)
        End If
    Wend
    Set readfrom = l
    delimiter = vbNullString
ElseIf isSequenceStop(token) Then
    Err.Raise 101, , "Unexpected " & token
Else
    Select Case firstchar(token)
        Case "#" 'reader macro invokation.
            If token = "#JSON" Then
                Set readfrom = readfrom(readJSON(tokens))
            Else
                Select Case tokens(1)
                    Case "(" '#( -> indicates anonymous function)
                        Set readfrom = list("error", "#(...) Anonymous function reader macro not implemented")
                    Case """"  '#"" -> indicates a regexp
                        Set readfrom = list("error", "#""..."" Regexp literal reader macro not implemented")
                    Case "{" '#{} -> indicates a set
                        tokens.Remove 1
                        Set tokens = prepend(tokens, "hash-set")
                        Set tokens = prepend(tokens, "{")
                        Set readfrom = readfrom(tokens)
                End Select
            End If
        Case "'"
            If tokens.count = 1 Then
                    Set tokens = prepend(tokens, restString(token))
                    Set readfrom = list("quote", readfrom(tokens))
            Else
                tokens.Remove 1
                If tokens(1) = "(" Then
                    tokens.Remove 1
                    Set tokens = prepend(tokens, "list")
                    Set tokens = prepend(tokens, "(")
                    Set readfrom = readfrom(tokens)
                Else
                    Set tokens = prepend(tokens, restString(token))
                    Set readfrom = list("quote", readfrom(tokens))
                End If
            End If
'        Case Chr(34) ' "
'            'tokens.Remove 1
'            Set tokens = prepend(tokens, ")")
'            Set tokens = prepend(tokens, "quote")
'            Set tokens = prepend(tokens, "(")
'            Set readfrom = readfrom(tokens)
        Case ":"
            'tokens.Remove 1
'            Set tokens = prepend(tokens, ")")
'            Set tokens = prepend(tokens, "quote")
'            Set tokens = prepend(tokens, token)
'            Set tokens = prepend(tokens, "(")
'            Set readfrom = readfrom(tokens)
            readfrom = token
        Case "`"
            tokens.Remove 1
            Set readfrom = list("quasiquote", readfrom(tokens))
        Case ","
            tokens.Remove 1
            If isQuasiSplice(token) Then
                Set readfrom = list("unquote-splicing", readfrom(tokens))
            Else
                Set readfrom = list("unquote", restString(token))
            End If
        Case Else
            readfrom = atom(token)
    End Select
End If
    
'public Object read() {
'  try {
'    Object token = nextToken();
'    if (token == "(")
'  return readTail(false);
'    else if (token == ")")
'  { warn("Extra ) ignored."); return read(); }
'    else if (token == ".")
'  { warn("Extra . ignored."); return read(); }
'    else if (token == "'")
'  return list("quote", read());
'    else if (token == "`")
'  return list("quasiquote", read());
'    else if (token == ",")
'  return list("unquote", read());
'    else if (token == ",@")
'  return list("unquote-splicing", read());
'    Else
'  return token;
'  } catch (IOException e) {
'    warn("On input, exception: " + e);
'    return EOF;
'  }
'}
       
End Function
'Converts a JSON string literal into an evaluation
'Mutates tokens in the process by consuming elements from the head.
'Returns '((string->json "somestring"), tokens)
Function readJSON(tokens As Collection) As Collection
Dim itm
Dim sb As StringBuilder
Set sb = New StringBuilder

sb.append Chr(34)

If tokens(1) = "#JSON" Then
    tokens.Remove 1
    While tokens(1) <> ""
        sb.append CStr(tokens(1))
        tokens.Remove 1
    Wend
End If
    
sb.append Chr(34)
    
'readJSON = sb.toString
Set readJSON = tokens
Set readJSON = prepend(tokens, ")")
Set readJSON = prepend(tokens, sb.toString)
Set readJSON = prepend(tokens, "read-JSON")
Set readJSON = prepend(tokens, "(")
Set sb = Nothing
    
End Function
Function getClosingDelim(ByRef delim As String) As String
Select Case delim
    Case "("
        getClosingDelim = ")"
    Case "["
        getClosingDelim = "]"
    Case "{"
        getClosingDelim = "}"
    Case Else
        Err.Raise 101, , "unknown delimiter " & delim
End Select

End Function
Function getSeqType(ByRef tok As String) As String
Select Case tok
    Case "("
        getSeqType = "list"
    Case "["
        getSeqType = "vector"
    Case "{"
        getSeqType = "hash-map"
    Case Else
        Err.Raise 101, , "Unknown sequence type " & tok
End Select

End Function
Function isSequenceStart(ByRef tok As String) As Boolean
Select Case tok
    Case "(", "[", "{"
        isSequenceStart = True
    Case Else
        isSequenceStart = False
End Select
    
End Function
Function isSequenceStop(ByRef tok As String, Optional ByRef tokstart As String) As Boolean
If tokstart = vbNullString Then
    Select Case tok
        Case ")", "]", "}"
            isSequenceStop = True
        Case Else
            isSequenceStop = False
    End Select
Else
    Select Case tokstart
        Case "("
            isSequenceStop = tok = ")"
        Case "["
            isSequenceStop = tok = "]"
        Case "{"
            isSequenceStop = tok = "}"
        Case Else
            Err.Raise 101, , "Unkown sequence delimiter " & tokstart
    End Select
End If

End Function

Function specialToken(ByRef tok As String) As Boolean
Select Case tok
    Case "'", ":", "`", "~", "~@"
        specialToken = True
End Select

End Function
Function firstchar(ByRef tok As String) As String
firstchar = Mid(tok, 1, 1)
End Function
Function nthString(n As Long, ByRef tok As String) As String
nthString = Mid(tok, n, 1)
End Function
Function restString(ByRef tok As String) As String
restString = Mid(tok, 2)
End Function
Function isQuasiSplice(ByRef tok As String) As Boolean
isQuasiSplice = Mid(tok, 1, 2) = "~@"
End Function
Function isKeyword(ByRef tok As String) As Boolean
isKeyword = Mid(tok, 1, 1) = ":"
End Function

Function whitespace(ByRef tok As String) As Boolean
Select Case tok
    Case "", vbCrLf, vbTab, " ", ",", Chr(10), Chr(13)
        whitespace = True
    Case Else
        whitespace = False
End Select
End Function

'def atom(token):
'    "Numbers become numbers; every other token is a symbol."
'    try: return int(token)
'    except ValueError:
'        try: return float(token)
'        except ValueError:
'            return Symbol(token)
'
Function atom(ByRef token As String, Optional stringlit As Boolean) As Variant
If stringlit Then
    atom = replace(token, Chr(34), "")
ElseIf val(token) <> 0 Then
    atom = CDbl(val(token))
Else
    Select Case token
        Case "0", "0.0", "0.00"
            atom = CDbl(val(token))
        Case Else
            atom = token
    End Select
End If
    
End Function
'def to_string(Exp):
'    "Convert a Python object back into a Lisp-readable string."
'    Return '('+' '.join(map(to_string, exp))+')' if isa(exp, list) else str(exp)
'
'def repl(prompt='lis.py> '):
'    "A prompt-read-eval-print loop."
'    While True:
'        val = eval(parse(raw_input(prompt)))
'        if val is not None: print to_string(val)

Public Function defPrim(env As Dictionary, ByRef name As String, id As Long, minArgs As Long, Optional maxArgs As Long) As Dictionary
Dim p As LispPrimitive
Set p = New LispPrimitive

p.name = name
p.idNumber = id
p.minArgs = minArgs
If maxArgs > 0 Then
    p.maxArgs = maxArgs
Else
    p.maxArgs = minArgs
End If
    
env.add name, p
Set defPrim = env
End Function
'TOM added 9 Nov 2012 -> allows us to use IFns, as defined in Lisp, from anywhere in vba.
'Maintains a function library.
Public Function getFunc(name As String) As LispPrimitive
With getPrimitiveFunctions()
    If .exists(name) Then
        Set getFunc = .item(name)
    Else
        Err.Raise 101, "Function " & name & " is not a known primitive function!"
    End If
End With

End Function
Public Function getPrimitiveFunctions() As Dictionary
If Lisp.primitiveFuncs Is Nothing Then _
    Set Lisp.primitiveFuncs = lispPrimitiveFunctions()
Set getPrimitiveFunctions = Lisp.primitiveFuncs
End Function
'TOM added 9 Nov 2012 -> adds a library of lisp functions, each is a useable IFn that takes a
'collection of args...so they can also be called from other VBA programs...
Public Function lispPrimitiveFunctions() As Dictionary
Dim defs As Collection
Dim itm
Dim fs As Collection
Dim prim As LispPrimitive

Dim n As Long
n = 999999999

Set defs = New Collection

For Each itm In list(list("+", lispPrimitives.lPLUS, 0, n), _
    list("map", lMAP, 1, n), list("list", lLIST, 0, n), _
    list("*", lTIMES, 0, n), _
    list("-", lMINUS, 1, n), _
    list("/", lDIVIDE, 1, n), _
    list("<", lLT, 2, n), _
    list("<=", lLE, 2, n), _
    list("=", lEQ, 2, n), _
    list(">", lGT, 2, n), _
    list(">=", lGE, 2, n), _
    list("abs", lABS, 1), _
    list("apply", lAPPLY, 2, n), _
    list("eval", lEVAL, 1, 2), _
    list("sin", lSIN, 1), _
    list("cos", lCOS, 1), _
    list("tan", lTAN, 1), _
    list("sqrt", lSQRT, 1), _
    list("modulo", lMODULO, 2), _
    list("length", lLENGTH, 1), _
    list("count", lLENGTH, 1), _
    list("vector", lVECTOR, 0, n), _
    list("hash-map", lHASHMAP, 0, n))
    defs.add itm
Next itm

    
For Each itm In list(list("hash-set", lHASHSET, 0, n), _
    list("even?", lEVENQ, 1, 1), _
    list("odd?", lODDQ, 1, 1), _
    list("number?", lNUMBERQ, 1, 1), _
    list("zero?", lZEROQ, 1, 1), _
    list("negative?", lNEGATIVEQ, 1, 1), _
    list("reduce", lREDUCE, 2, n), _
    list("read-JSON", lREADJSON, 1, 1), _
    list("write-JSON", lWRITEJSON, 1, 1))
    defs.add itm
Next itm

'Sequence functions....
For Each itm In list(list("take", lTAKE, 2), _
    list("take-while", lTAKEWHILE, 2), _
    list("iterate", lITERATE, 2), _
    list("drop", lDROP, 2), _
    list("first", lFIRST, 1), _
    list("rest", lREST, 1), _
    list("next", lNEXT, 1), _
    list("lazy-cons", lLAZYCONS, 2), _
    list("seq", lSEQ, 2), _
    list("filter", lFILTER, 2), _
    list("concat", lCONCAT, 2), _
    list("last", lLAST, 2), _
    list("but-last", lBUTLAST, 2), _
    list("nth", lNTH, 2), _
    list("inc", lINC, 1), _
    list("dec", lDEC, 1), _
    list("repeat", lREPEAT, 1), _
    list("repeatedly", lREPEATEDLY, 1), _
    list("second", lSECOND, 1))
    
defs.add itm
Next itm

'Sequence functions....
For Each itm In list(list("rand", lRAND, 0, 1), _
    list("rand-int", lRANDINT, 1), _
    list("rand-between", lRANDBETWEEN, 2), _
    list("load-file", lLOADFILE, 1), _
    list("range", lrange, 1), _
    list("reverse", lreverse, 1), _
    list("sort-by", lsortby, 2), _
    list("sort", lsort, 1, 2), _
    list("conj", lconj, 2))
defs.add itm
Next itm


Dim coll As Collection
Set lispPrimitiveFunctions = New Dictionary
For Each coll In defs
    If coll.count > 3 Then
        Set prim = makePrimitive(CLng(coll(2)), CLng(coll(3)), CLng(coll(4)))
    Else
        Set prim = makePrimitive(CLng(coll(2)), CLng(coll(3)))
    End If
    lispPrimitiveFunctions.add coll(1), prim
Next coll
    
End Function
Public Function asFunc(x As Variant) As IFn
Set asFunc = x
End Function


Public Function installPrimitives(env As Dictionary) As Dictionary

Set installPrimitives = DictionaryLib.mergeDicts(env, getPrimitiveFunctions())
env.add "pi", lispPI
env.add "e", lispE
'Set env = defPrim(env, "+", lispPrimitives.lPLUS, 0, n)
'Set env = defPrim(env, "map", lMAP, 1, n)
'Set env = defPrim(env, "list", lLIST, 0, n)
'Set env = defPrim(env, "*", lTIMES, 0, n)
'Set env = defPrim(env, "-", lMINUS, 1, n)
'Set env = defPrim(env, "/", lDIVIDE, 1, n)
'Set env = defPrim(env, "<", lLT, 2, n)
'Set env = defPrim(env, "<=", lLE, 2, n)
'Set env = defPrim(env, "=", lEQ, 2, n)
'Set env = defPrim(env, ">", lGT, 2, n)
'Set env = defPrim(env, ">=", lGE, 2, n)
'Set env = defPrim(env, "abs", lABS, 1)
'Set env = defPrim(env, "apply", lAPPLY, 2, n)
'Set env = defPrim(env, "eval", lEVAL, 1, 2)
'Set env = defPrim(env, "sin", lSIN, 1)
'Set env = defPrim(env, "cos", lCOS, 1)
'Set env = defPrim(env, "tan", lTAN, 1)
'Set env = defPrim(env, "sqrt", lSQRT, 1)
'Set env = defPrim(env, "modulo", lMODULO, 2)
'Set env = defPrim(env, "length", lLENGTH, 1)
'Set env = defPrim(env, "count", lLENGTH, 1)
'Set env = defPrim(env, "vector", lVECTOR, 0, n)
'Set env = defPrim(env, "hash-map", lHASHMAP, 0, n)
'Set env = defPrim(env, "hash-set", lHASHSET, 0, n)
'Set env = defPrim(env, "even?", lEVENQ, 1, 1)
'Set env = defPrim(env, "odd?", lODDQ, 1, 1)
'Set env = defPrim(env, "number?", lNUMBERQ, 1, 1)
'Set env = defPrim(env, "zero?", lZEROQ, 1, 1)
'Set env = defPrim(env, "negative?", lNEGATIVEQ, 1, 1)
'
''reduce
'Set env = defPrim(env, "reduce", lREDUCE, 2, n)
'
''Literal JSON support
'Set env = defPrim(env, "read-JSON", lREADJSON, 1, 1)
'Set env = defPrim(env, "write-JSON", lWRITEJSON, 1, 1)
'
''Sequence functions....
'Set env = defPrim(env, "take", lTAKE, 2)
'Set env = defPrim(env, "take-while", lTAKEWHILE, 2)
'Set env = defPrim(env, "iterate", lITERATE, 2)
'Set env = defPrim(env, "drop", lDROP, 2)
'Set env = defPrim(env, "first", lFIRST, 1)
'Set env = defPrim(env, "rest", lREST, 1)
'Set env = defPrim(env, "next", lNEXT, 1)
'Set env = defPrim(env, "lazy-cons", lLAZYCONS, 2)
'Set env = defPrim(env, "seq", lSEQ, 2)
'Set env = defPrim(env, "filter", lFILTER, 2)
'Set env = defPrim(env, "concat", lCONCAT, 2)
'Set env = defPrim(env, "last", lLAST, 2)
'Set env = defPrim(env, "but-last", lBUTLAST, 2)
'Set env = defPrim(env, "nth", lNTH, 2)
'
'Set env = defPrim(env, "inc", lINC, 1)
'Set env = defPrim(env, "dec", lDEC, 1)



'    env
'     .defPrim("*",          TIMES,     0, n)
'     .defPrim("*",          TIMES,     0, n)
'     .defPrim("+",          PLUS,      0, n)
'     .defPrim("-",          MINUS,     1, n)
'     .defPrim("/",          DIVIDE,    1, n)
'     .defPrim("<",          LT,        2, n)
'     .defPrim("<=",         LE,        2, n)
'     .defPrim("=",          EQ,        2, n)
'     .defPrim(">",          GT,        2, n)
'     .defPrim(">=",         GE,        2, n)
'     .defPrim("abs",        ABS,       1)
'     .defPrim("acos",       ACOS,      1)
'     .defPrim("append",         APPEND,    0, n)
'     .defPrim("apply",      APPLY,     2, n)
'     .defPrim("asin",       ASIN,      1)
'     .defPrim("assoc",      ASSOC,     2)
'     .defPrim("assq",       ASSQ,      2)
'     .defPrim("assv",       ASSV,      2)
'     .defPrim("atan",       ATAN,      1)
'     .defPrim("boolean?",   BOOLEANQ,  1)
'     .defPrim("caaaar",         CXR,       1)
'     .defPrim("caaadr",         CXR,       1)
'     .defPrim("caaar",          CXR,       1)
'     .defPrim("caadar",         CXR,       1)
'     .defPrim("caaddr",         CXR,       1)
'     .defPrim("caadr",          CXR,       1)
'     .defPrim("caar",           CXR,       1)
'     .defPrim("cadaar",         CXR,       1)
'     .defPrim("cadadr",         CXR,       1)
'     .defPrim("cadar",          CXR,       1)
'     .defPrim("caddar",         CXR,       1)
'     .defPrim("cadddr",         CXR,       1)
'     .defPrim("caddr",      THIRD,     1)
'     .defPrim("cadr",           SECOND,    1)
'     .defPrim("call-with-current-continuation",        CALLCC,    1)
'     .defPrim("call-with-input-file", CALLWITHINPUTFILE, 2)
'     .defPrim("call-with-output-file", CALLWITHOUTPUTFILE, 2)
'     .defPrim("car",        CAR,       1)
'     .defPrim("cdaaar",         CXR,       1)
'     .defPrim("cdaadr",         CXR,       1)
'     .defPrim("cdaar",          CXR,       1)
'     .defPrim("cdadar",         CXR,       1)
'     .defPrim("cdaddr",         CXR,       1)
'     .defPrim("cdadr",          CXR,       1)
'     .defPrim("cdar",           CXR,       1)
'     .defPrim("cddaar",         CXR,       1)
'     .defPrim("cddadr",         CXR,       1)
'     .defPrim("cddar",          CXR,       1)
'     .defPrim("cdddar",         CXR,       1)
'     .defPrim("cddddr",         CXR,       1)
'     .defPrim("cdddr",          CXR,       1)
'     .defPrim("cddr",           CXR,       1)
'     .defPrim("cdr",        CDR,       1)
'     .defPrim("char->integer",  CHARTOINTEGER,      1)
'     .defPrim("char-alphabetic?",CHARALPHABETICQ,      1)
'     .defPrim("char-ci<=?",     CHARCICMP+LE, 2)
'     .defPrim("char-ci<?" ,     CHARCICMP+LT, 2)
'     .defPrim("char-ci=?" ,     CHARCICMP+EQ, 2)
'     .defPrim("char-ci>=?",     CHARCICMP+GE, 2)
'     .defPrim("char-ci>?" ,     CHARCICMP+GT, 2)
'     .defPrim("char-downcase",  CHARDOWNCASE,      1)
'     .defPrim("char-lower-case?",CHARLOWERCASEQ,      1)
'     .defPrim("char-numeric?",  CHARNUMERICQ,      1)
'     .defPrim("char-upcase",    CHARUPCASE,      1)
'     .defPrim("char-upper-case?",CHARUPPERCASEQ,      1)
'     .defPrim("char-whitespace?",CHARWHITESPACEQ,      1)
'     .defPrim("char<=?",        CHARCMP+LE, 2)
'     .defPrim("char<?",         CHARCMP+LT, 2)
'     .defPrim("char=?",         CHARCMP+EQ, 2)
'     .defPrim("char>=?",        CHARCMP+GE, 2)
'     .defPrim("char>?",         CHARCMP+GT, 2)
'     .defPrim("char?",      CHARQ,     1)
'     .defPrim("close-input-port", CLOSEINPUTPORT, 1)
'     .defPrim("close-output-port", CLOSEOUTPUTPORT, 1)
'     .defPrim("complex?",   NUMBERQ,   1)
'     .defPrim("cons",       CONS,      2)
'     .defPrim("cos",        COS,       1)
'     .defPrim("current-input-port", CURRENTINPUTPORT, 0)
'     .defPrim("current-output-port", CURRENTOUTPUTPORT, 0)
'     .defPrim("display",        DISPLAY,   1, 2)
'     .defPrim("eof-object?",    EOFOBJECTQ, 1)
'     .defPrim("eq?",        EQQ,       2)
'     .defPrim("equal?",     EQUALQ,    2)
'     .defPrim("eqv?",       EQVQ,      2)
'     .defPrim("eval",           EVAL,      1, 2)
'     .defPrim("even?",          EVENQ,     1)
'     .defPrim("exact?",         INTEGERQ,  1)
'     .defPrim("exp",        EXP,       1)
'     .defPrim("expt",       EXPT,      2)
'     .defPrim("force",          FORCE,     1)
'     .defPrim("for-each",       FOREACH,   1, n)
'     .defPrim("gcd",            GCD,       0, n)
'     .defPrim("inexact?",       INEXACTQ,  1)
'     .defPrim("input-port?",    INPUTPORTQ, 1)
'     .defPrim("integer->char",  INTEGERTOCHAR,      1)
'     .defPrim("integer?",       INTEGERQ,  1)
'     .defPrim("lcm",            LCM,       0, n)
'     .defPrim("length",     LENGTH,    1)
'     .defPrim("list",       LIST,      0, n)
'     .defPrim("list->string",   LISTTOSTRING, 1)
'     .defPrim("list->vector",   LISTTOVECTOR,      1)
'     .defPrim("list-ref",   LISTREF,   2)
'     .defPrim("list-tail",  LISTTAIL,  2)
'     .defPrim("list?",          LISTQ,     1)
'     .defPrim("load",           LOAD,      1)
'     .defPrim("log",        LOG,       1)
'     .defPrim("macro-expand",   MACROEXPAND,1)
'     .defPrim("make-string",    MAKESTRING,1, 2)
'     .defPrim("make-vector",    MAKEVECTOR,1, 2)
'     .defPrim("map",            MAP,       1, n)
'     .defPrim("max",        MAX,       1, n)
'     .defPrim("member",     MEMBER,    2)
'     .defPrim("memq",       MEMQ,      2)
'     .defPrim("memv",       MEMV,      2)
'     .defPrim("min",        MIN,       1, n)
'     .defPrim("modulo",         MODULO,    2)
'     .defPrim("negative?",      NEGATIVEQ, 1)
'     .defPrim("newline",    NEWLINE,   0, 1)
'     .defPrim("not",        NOT,       1)
'     .defPrim("null?",      NULLQ,     1)
'     .defPrim("number->string", NUMBERTOSTRING,   1, 2)
'     .defPrim("number?",    NUMBERQ,   1)
'     .defPrim("odd?",           ODDQ,      1)
'     .defPrim("open-input-file",OPENINPUTFILE, 1)
'     .defPrim("open-output-file", OPENOUTPUTFILE, 1)
'     .defPrim("output-port?",   OUTPUTPORTQ, 1)
'     .defPrim("pair?",      PAIRQ,     1)
'     .defPrim("peek-char",      PEEKCHAR,  0, 1)
'     .defPrim("positive?",      POSITIVEQ, 1)
'     .defPrim("procedure?",     PROCEDUREQ,1)
'     .defPrim("quotient",       QUOTIENT,  2)
'     .defPrim("rational?",      INTEGERQ, 1)
'     .defPrim("read",       READ,      0, 1)
'     .defPrim("read-char",      READCHAR,  0, 1)
'     .defPrim("real?",          NUMBERQ,   1)
'     .defPrim("remainder",      REMAINDER, 2)
'     .defPrim("reverse",    REVERSE,   1)
'     .defPrim("round",      ROUND,     1)
'     .defPrim("set-car!",   SETCAR,    2)
'     .defPrim("set-cdr!",   SETCDR,    2)
'     .defPrim("sin",        SIN,       1)
'     .defPrim("sqrt",       SQRT,      1)
'     .defPrim("string",     STRING,    0, n)
'     .defPrim("string->list",   STRINGTOLIST, 1)
'     .defPrim("string->number", STRINGTONUMBER,   1, 2)
'     .defPrim("string->symbol", STRINGTOSYMBOL,   1)
'     .defPrim("string-append",  STRINGAPPEND, 0, n)
'     .defPrim("string-ci<=?",   STRINGCICMP+LE, 2)
'     .defPrim("string-ci<?" ,   STRINGCICMP+LT, 2)
'     .defPrim("string-ci=?" ,   STRINGCICMP+EQ, 2)
'     .defPrim("string-ci>=?",   STRINGCICMP+GE, 2)
'     .defPrim("string-ci>?" ,   STRINGCICMP+GT, 2)
'     .defPrim("string-length",  STRINGLENGTH,   1)
'     .defPrim("string-ref",     STRINGREF, 2)
'     .defPrim("string-set!",    STRINGSET, 3)
'     .defPrim("string<=?",      STRINGCMP+LE, 2)
'     .defPrim("string<?",       STRINGCMP+LT, 2)
'     .defPrim("string=?",       STRINGCMP+EQ, 2)
'     .defPrim("string>=?",      STRINGCMP+GE, 2)
'     .defPrim("string>?",       STRINGCMP+GT, 2)
'     .defPrim("string?",    STRINGQ,   1)
'     .defPrim("substring",  SUBSTRING, 3)
'     .defPrim("symbol->string", SYMBOLTOSTRING,   1)
'     .defPrim("symbol?",    SYMBOLQ,   1)
'     .defPrim("tan",        TAN,       1)
'     .defPrim("vector",     VECTOR,    0, n)
'     .defPrim("vector->list",   VECTORTOLIST, 1)
'     .defPrim("vector-length",  VECTORLENGTH, 1)
'     .defPrim("vector-ref",     VECTORREF, 2)
'     .defPrim("vector-set!",    VECTORSET, 3)
'     .defPrim("vector?",        VECTORQ,   1)
'     .defPrim("write",      WRITE,     1, 2)
'     .defPrim("write-char",     DISPLAY,   1, 2)
'     .defPrim("zero?",          ZEROQ,     1)
'
'     ///////////// Extensions ////////////////
'
'     .defPrim("new",            NEW,       1)
'     .defPrim("class",          CLASS,     1)
'     .defPrim("method",         METHOD,    2, n)
'     .defPrim("exit",           EXIT,      0, 1)
'     .defPrim("error",          ERROR,     0, n)
'     .defPrim("time-call",          TIMECALL,  1, 2)
'     .defPrim("_list*",             LISTSTAR,  0, n)
'       ;
'
'     return env;
'  }
End Function
'Installs additional libraries, where libarires lisp expressions contained in libs.
Public Function installLibraries(env As Dictionary, Optional libs As Collection) As Dictionary
Dim expr
Set installLibraries = env

If exists(libs) Then
    For Each expr In libs
        eval expr, env
    Next expr
End If

End Function

'A quick test of the Lisp (more scheme) interpreter in VBA......I Don't know if this is an accomplishment
'or an attrocity, but it's kept me interested!  This is running from within Excel, in the VBA IDE.
Sub tst()
Dim env As Dictionary
Set env = installPrimitives(makeEnvironment)

pprint read("(+ 2 3)")
pprint eval(read("(+ 2 3)"), env)
pprint eval(read("(map (lambda (x) (+ 1 x)) (list 1 2 3))"), env)
pprint eval(read( _
    "(begin (define pi 3.14) " & _
           "(define area (lambda (r) (* r r pi)))" & _
           "(map area (list 1 2 3 4)))"), env)

pprint read("2")

End Sub

Sub tstbind()
Dim x As Variant
Dim y As Variant


Set y = list(1, 2, 3)
bind x, y

End Sub
'define lambda's inline in VBA
Function lambda(params As Collection, body As String, Optional env As Dictionary) As IFn

If env Is Nothing Then Set env = Lisp.getGlobalLispEnv
Set lambda = makeClosure(params, read(body), env)

End Function

'FROM SICP, separating execution from analysis.
Function analyze(expr As Variant) As Variant
Select Case expType(expr)
    Case expressionTypes.selfEvaluatingQ
        'bind analyze, list("lambda", list("env"), expr)
    Case expressionTypes.quotedQ
        'not implemented....
        'Dim txt As String
        'txt = textOfQuotation(expr)
        'bind analyze, lambda(list("env"), txt)
    Case expressionTypes.variableQ
        'bind analyze, lambda(list("env"), "(lookup-variable-value " & expr & " env)")
    Case expressionTypes.assignmentQ
    Case expressionTypes.definitionQ
    Case expressionTypes.ifQ
    Case expressionTypes.lambdaQ
    Case expressionTypes.beginQ
    Case expressionTypes.condQ
    Case expressionTypes.applicationQ
    Case expressionTypes.unknown
End Select
    
End Function
Function textOfQuotation(expr As Collection) As String
textOfQuotation = printstr(restList(expr))
End Function
Function expType(expr As Variant) As expressionTypes
End Function

Function apply(fn As IFn, args As Collection) As Variant
bind apply, fn.apply(args)
End Function


'Calling anonymous functions from VBA
Sub lambdatest()
Dim i As Long
Dim add2  As IFn

Set add2 = lambda(list("x"), "(* 2 x)")

pprint map(add2, floatList(33.5, 4.4, -2))
pprint reduce(lambda(list("acc", "x"), "(+ acc x)"), 0#, floatList(10))

End Sub

'For each element in the sequence Args, apply f to the element, accumulating a
'collection of results.  Returns a 1:1 mapping of function applications.
Public Function map(f As IFn, ByRef args As Variant) As Variant
Dim acc As Collection
Dim key
Dim c As Collection
Dim d As Dictionary
Dim s As ISeq
Dim i As Long
Set acc = New Collection

Select Case TypeName(args)
    Case "Collection"
        Set c = asCollection(args)
        For i = 1 To args.count
            acc.add f.apply(list(args(i)))
        Next i
    Case "Dictionary"
        Set d = asDict(args)
        For Each key In d
            bind acc, f.apply(list(acc, list(key, d(key))))
        Next key
    Case Else
        If isSeq(args) Then
            Set s = args
            Set map = SeqLib.seqMap(f, s)
            Set s = Nothing
            Exit Function
        Else
            Err.Raise 101, , "Do not know how to reduce " & TypeName(args)
        End If
End Select


Set map = acc
End Function

'Takes a function f, whose first arg is an accumulator, and second arg is an item to
'be reduced.
'For each element in the sequence Args, apply f to the acc and the element, accumulating
'results in acc.
Public Function reduce(f As IFn, Optional initialval As Variant, Optional ByRef args) As Variant
Dim i As Long
Dim offset As Long
Dim key
Dim d As Dictionary
Dim c As Collection
Dim s As ISeq

Dim acc As Variant

If args Is Nothing Then
    Err.Raise 101, , "No arguements provided for reduction.  Context is undefined"
ElseIf IsMissing(initialval) Then
    bind initialval, args(1)
    offset = 1
End If

bind acc, initialval

Select Case TypeName(args)
    Case "Collection"
        Set c = asCollection(args)
        For i = 1 + offset To c.count
            bind acc, f.apply(list(acc, c(i)))
        Next i
    Case "Dictionary"
        Set d = asDict(args)
        For Each key In d
            bind acc, f.apply(list(acc, list(key, d(key))))
        Next key
    Case Else
        If isSeq(args) Then
            Set s = args
            If Not nil(s.fst) Then
                bind acc, f.apply(list(acc, s.fst))
                While exists(s.more)
                    Set s = s.more
                    bind acc, f.apply(list(acc, s.fst))
                Wend
            End If
            Set s = Nothing
        Else
            Err.Raise 101, , "Do not know how to reduce " & TypeName(args)
        End If
End Select
            
            
bind reduce, acc
clearbind acc
End Function

'Simulates a simple read-evaluate-print composition, without Looping.
'Facilitates evaluating lisp Expressions from the immediate window.
Public Sub rep(ByRef expression As String, Optional env As Dictionary)
pprint eval(read(expression), env)
'resetLisp
End Sub
Function reval(ByRef expression As String, Optional env As Dictionary) As Variant
bind reval, eval(read(expression))
End Function
'Fetch a varname from bindings in the lisp Environment.  Obeys lexical scoping rules
'using findOuter.  Typically used to lookup global vars from the lisp environment.
Function resolve(ByRef varname As String, Optional env As Dictionary) As Variant
If env Is Nothing Then Set env = Lisp.getGlobalLispEnv
bind resolve, findOuter(env, varname)
End Function


'Reads a file...depending on the extension (i.e. if it's JSON), it'll automatically
'interpret it as a JSON-like file, otherwise uses the Lisp evaluator (with clojure syntax
'extensions for datastructures) to evaluate the data...
Public Function readFile(path As String) As Variant
Dim fl As String
Dim frm
If InStr(1, UCase(path), ".JSON") = 0 Then 'normal read...
    fl = SerializationLib.readString(path)
    fl = "(" & fl & ")"
    For Each frm In read(fl)
        eval frm
    Next frm
Else
    fl = "#JSON" & SerializationLib.readString(path)
    bind readFile, reval(fl)
End If



End Function
