Option Explicit
    
Public minArgs As Long
Public maxArgs As Long
Public idNumber As Long
Public name As String

Public Enum lispPrimitives
   lEQ = 0
   lLT = 1
   lGT = 2
   lGE = 3
   lLE = 4
   lABS = 5
   lEOF_OBJECT = 6
   lEQQ = 7
   lEQUALQ = 8
   lFORCE = 9
   lCAR = 10
   lFLOOR = 11
   lCeiling = 12
   lCons = 13
   lDIVIDE = 14
   lLENGTH = 15
   lLIST = 16
   lLISTQ = 17
   lAPPLY = 18
   lMAX = 19
   lMIN = 20
   lMINUS = 21
   lNEWLINE = 22
   lNOT = 23
   lNULLQ = 24
   lNUMBERQ = 25
   lPAIRQ = 26
   lPLUS = 27
   lPROCEDUREQ = 28
   lREAD = 29
   lCDR = 30
   lROUND = 31
   lSECOND = 32
   lSYMBOLQ = 33
   lTIMES = 34
   lTRUNCATE = 35
   lWRITE = 36
   lAPPEND = 37
   lBOOLEANQ = 38
   lSQRT = 39
   lEXPT = 40
   lreverse = 41
   lASSOC = 42
   lASSQ = 43
   lASSV = 44
   lMEMBER = 45
   lMEMQ = 46
   lMEMV = 47
   lEQVQ = 48
   lLISTREF = 49
   lLISTTAIL = 50
   lSTRINQ = 51
   lMAKESTRING = 52
   lSTRING = 53
   lSTRINGLENGTH = 54
   lSTRINGREF = 55
   lSTRINGSET = 56
   lSUBSTRING = 57
   lSTRINGAPPEND = 58
   lSTRINGTOLIST = 59
   lLISTTOSTRING = 60
   lSYMBOLTOSTRING = 61
   lSTRINGTOSYMBOL = 62
   lEXP = 63
   lLOG = 64
   lSIN = 65
   lCOS = 66
   lTAN = 67
   lACOS = 68
   lASIN = 69
   lATAN = 70
   lNUMBERTOSTRING = 71
   lSTRINGTONUMBER = 72
   lCHARQ = 73
   lCHARALPHABETICQ = 74
   lCHARNUMERICQ = 75
   lCHARWHITESPACEQ = 76
   lCHARUPPERCASEQ = 77
   lCHARLOWERCASEQ = 78
   lCHARTOINTEGER = 79
   lINTEGERTOCHAR = 80
   lCHARUPCASE = 81
   lCHARDOWNCASE = 82
   lSTRINGQ = 83
   lVECTORQ = 84
   lMAKEVECTOR = 85
   lVECTOR = 86
   lVECTORLENGTH = 87
   lVECTORREF = 88
   lVECTORSET = 89
   lLISTTOVECTOR = 90
   lMAP = 91
   lFOREACH = 92
   lCALLCC = 93
   lVECTORTOLIST = 94
   lLOAD = 95
   lDISPLAY = 96
   lINPUTPORTQ = 98
   lCURRENTINPUTPORT = 99
   lOPENINPUTFILE = 100
   lCLOSEINPUTPORT = 101
   lOUTPUTPORTQ = 103
   lCURRENTOUTPUTPORT = 104
   lOPENOUTPUTFILE = 105
   lCLOSEOUTPUTPORT = 106
   lREADCHAR = 107
   lPEEKCHAR = 108
   lEVAL = 109
   lQUOTIENT = 110
   lREMAINDER = 111
   lMODULO = 112
   lTHIRD = 113
   lEOFOBJECTQ = 114
   lGCD = 115
   lLCM = 116
   lCXR = 117
   lODDQ = 118
   lEVENQ = 119
   lZEROQ = 120
   lPOSITIVEQ = 121
   lNEGATIVEQ = 122
   lCHARCMP = 123 '/* to 127 */
   lCHARCICMP = 128 '/* to 132 */,
   lSTRINGCMP = 133 '/* to 137 */
   lSTRINGCICMP = 138 '/* to 142 */,
   lEXACTQ = 143
   lINEXACTQ = 144
   lINTEGERQ = 145
   lCALLWITHINPUTFILE = 146
   lCALLWITHOUTPUTFILE = 147
   lNEW = -1
   lCLASS = -2
   lMETHOD = -3
   lEXIT = -4
   lSETCAR = -5
   lSETCDR = -6
   lTIMECALL = -11
   lMACROEXPAND = -12
   lERROR = -13
   lLISTSTART = -14
   'tom stuff
   lREDUCE = -15
   lHASHSET = -16
   lHASHMAP = -17
   lTAKE = -18
   lTAKEWHILE = -19
   lITERATE = -20
   lDROP = -21
   lFIRST = -22
   lREST = -23
   lNEXT = -24
   lLAZYCONS = -25
   lSEQ = -26
   lLAZYSEQ = -27
   lREPEAT = -277
   lREPEATEDLY = -288
   lREADJSON = -28
   lWRITEJSON = -29
   lFILTER = -30
   lCONCAT = -31
   lLAST = -32
   lBUTLAST = -33
   lNTH = -34
   lINC = -35
   lDEC = -36
   lconj = -37
   lRAND = -38
   lRANDINT = -39
   lRANDBETWEEN = -40
   lKEYS = -41
   lVALS = -42
   lGET = -43  'polymorphic function for getting elements from a collection
   lLOADFILE = -44
   lrange = -45
   lsortby = -47
   lsort = -48
End Enum
Implements IFn

   
'  //////////////// Extensions ////////////////
'
'    static final int NEW = -1, CLASS = -2, METHOD = -3, EXIT = -4,
'      SETCAR = -5, SETCDR = -6, TIMECALL = -11, MACROEXPAND = -12,
'      ERROR = -13, LISTSTAR = -14
'    ;

'/** Apply a primitive to a list of arguments. **/
Public Function apply(args As Collection) As Variant
'  //First make sure there are the right number of arguments.
Dim nargs As Long

nargs = args.count
If (nargs < minArgs) Then
    Err.Raise "too few args, " + nargs + ", for " + name + ": " + printstr(args)
ElseIf (nargs > maxArgs) Then
    Err.Raise "too many args, " + nargs + ", for " + name + ": " + printstr(args)
End If

Dim x As Variant
Dim y As Variant

If args.count > 0 Then
        bind x, args(1)
    If args.count > 1 Then
        bind y, args(2)
    End If
End If

Select Case idNumber
    
    '  ////////////////  SECTION 6.1 BOOLEANS
    Case lNOT
        apply = truth(x = False)
    Case lBOOLEANQ
        apply = truth(x = True Or x = False)
    '  ////////////////  SECTION 6.2 EQUIVALENCE PREDICATES
    Case lEQVQ
        'apply = truth(eqv(x, y))
    Case lEQQ
        apply = truth(x = y)
    Case lEQUALQ
        'apply = truth(equal(x, y))
        '
    '////////////////  SECTION 6.3 LISTS AND PAIRS
    Case lPAIRQ
        apply = asCollection(x).count = 2
    Case lLISTQ
        apply = isa(x, "Collection")
'    case CXR:           for (int i = name.length()-2; i >= 1; i--)
'                          x = (name.charAt(i) == 'a') ? first(x) : rest(x);
'                        return x;
'    case CONS:      return cons(x, y);
    Case lCAR
        bind apply, x
    Case lCDR
        'apply = rest(x)
'    case SETCAR:        return setFirst(x, y);
'    case SETCDR:        return setRest(x, y);
    Case lSECOND
        bind apply, SeqLib.nth(1, seq(x))
    Case lTHIRD:
        bind apply, SeqLib.nth(2, seq(x))
    Case lNULLQ
        apply = truth(IsNull(x))
    Case lLIST
        Set apply = args
    Case lLENGTH
        apply = length(x)
    Case lAPPEND
        'args == null) ? null : append(args);
        
    Case lreverse
        Set apply = SeqLib.seqReverse(args)

'    case LISTTAIL:  for (int k = (int)num(y); k>0; k--) x = rest(x);
'                        return x;
'    case LISTREF:   for (int k = (int)num(y); k>0; k--) x = rest(x);
'                        return first(x);
'    case MEMQ:          return memberAssoc(x, y, 'm', 'q');
'    case MEMV:          return memberAssoc(x, y, 'm', 'v');
'    case MEMBER:        return memberAssoc(x, y, 'm', ' ');
'    case ASSQ:          return memberAssoc(x, y, 'a', 'q');
'    case ASSV:          return memberAssoc(x, y, 'a', 'v');
'    case ASSOC:         return memberAssoc(x, y, 'a', ' ');

'      ////////////////  SECTION 6.4 SYMBOLS
    Case lSYMBOLQ
        apply = isa(x, "String")
    Case lSYMBOLTOSTRING
        apply = CStr(x) 'sym(x).toCharArray();
    Case lSTRINGTOSYMBOL
        'return new String(str(x)).intern();
'      ////////////////  SECTION 6.5 NUMBERS
    Case lNUMBERQ
        apply = IsNumeric(x)
    Case lODDQ
        apply = Not (Abs(x Mod 2) = 0)
    Case lEVENQ
        apply = Abs(x Mod 2) = 0
    Case lZEROQ
        apply = x = 0
    Case lPOSITIVEQ
        apply = x > 0
    Case lNEGATIVEQ
        apply = x < 0
'    case INTEGERQ:      return truth(isExact(x));
'    case INEXACTQ:      return truth(!isExact(x));
    Case lLT
        apply = x < y
    Case lGT
        apply = x > y
    Case lEQ
        apply = x = y
    Case lLE
        apply = x <= y
    Case lGE
        apply = x >= y
    Case lMAX
        apply = numcompute(args, "X", CSng(x))
    Case lMIN
        apply = numcompute(args, "N", CSng(x))
    Case lPLUS
        apply = numcompute(args, "+", 0#)
    Case lMINUS:
        apply = numcompute(args, "-", 0)
    Case lTIMES
        apply = numcompute(args, "*", 1#)
    Case lDIVIDE
        apply = numcompute(args, "/", 0)
'    case QUOTIENT:      double d = num(x)/num(y);
'                        return num(d > 0 ? Math.floor(d) : Math.ceil(d));
'    case REMAINDER:     return num((long)num(x) % (long)num(y));
'    case MODULO:        long xi = (long)num(x), yi = (long)num(y), m = xi % yi;
'                        return num((xi*yi > 0 || m == 0) ? m : m + yi);
    Case lABS
        apply = Abs(x)
    Case lFLOOR
        'apply = Floor(x)
'    case CEILING:   return num(Math.ceil(num(x)));
'    case TRUNCATE:  d = num(x);
'                        return num((d < 0.0) ? Math.ceil(d) : Math.floor(d));
'    case ROUND:     return num(Math.round(num(x)));
    Case lEXP
        apply = exp(CDbl(x))
    Case lLOG
        apply = log(CDbl(x))
    Case lSIN
        apply = Sin(CDbl(x))
    Case lCOS
        apply = Cos(CDbl(x))
    Case lTAN
        apply = Tan(CDbl(x))
'    case ASIN:          return num(Math.asin(num(x)));
'    case ACOS:          return num(Math.acos(num(x)));
'    case ATAN:          return num(Math.atan(num(x)));
    Case lSQRT
        apply = x ^ 0.5
    Case lEXPT
        apply = x ^ y
'    case NUMBERTOSTRING:return numberToString(x, y);
'    case STRINGTONUMBER:return stringToNumber(x, y);
'    case GCD:           return (args == null) ? ZERO : gcd(args);
'    case LCM:           return (args == null) ? ONE  : lcm(args);


'      ////////////////  SECTION 6.9 CONTROL FEATURES
'    case EVAL:          return interp.eval(x);
'    case FORCE:         return (!(x instanceof Procedure)) ? x
'              : proc(x).apply(interp, null);
    Case lMACROEXPAND
'        bind apply, Macro.macroExpand(interp, x)
    Case lPROCEDUREQ
        apply = truth(TypeName(x) = "IFn")
    Case lAPPLY
        bind apply, Lisp.apply(asFunc(x), restList(args))
    Case lMAP
        bind apply, Lisp.map(asFunc(x), seq(y))
    Case lREDUCE
        If args.count = 3 Then
            bind apply, Lisp.reduce(asFunc(x), y, args(3))
        ElseIf args.count = 2 Then
            bind apply, Lisp.reduce(asFunc(x), , asCollection(y))
        Else
            Err.Raise 101, , "Invalid number of args to reduce!"
        End If
'    case FOREACH:       return map(proc(x), rest(args), interp, null);
'    case CALLCC:        RuntimeException cc = new RuntimeException();
'                        Continuation proc = new Continuation(cc);
'                    try { return proc(x).apply(interp, list(proc)); }
'            catch (RuntimeException e) {
'                if (e == cc) return proc.value; else throw e;
    Case lVECTOR
        bind apply, listToVector(args)
    Case lHASHSET
        bind apply, SetLib.setOfList(args)
    Case lHASHMAP
        bind apply, DictionaryLib.dictOfAList(args)
    Case lREADJSON
        Dim jstring As String
        jstring = CStr(x)
        Select Case firstchar(jstring)
            Case "{"
                bind apply, SerializationLib.JSONtoDictionary(jstring)
            Case "["
                bind apply, SerializationLib.JSONtoCollection(jstring)
            Case Else
                bind apply, SerializationLib.JSONtoPrimitive(jstring)
        End Select
    Case lWRITEJSON
        bind apply, SerializationLib.jParse(x)
    'Sequence operations.
    Case lTAKE
        bind apply, SeqLib.take(CLng(x), seq(y))
    Case lTAKEWHILE
        bind apply, SeqLib.takewhile(asFunc(x), seq(y))
    Case lITERATE
        bind apply, SeqLib.iterate(asFunc(x), y)
    Case lDROP
        bind apply, SeqLib.drop(CLng(x), seq(y))
    Case lFIRST
        bind apply, SeqLib.first(seq(x))
    Case lREST
        bind apply, SeqLib.rest(seq(x))
    Case lNEXT
        bind apply, SeqLib.seqNext(seq(x))
    Case lLAZYCONS
        bind apply, SeqLib.lazyCons(x, asFunc(y))
    Case lSEQ
        bind apply, SeqLib.seq(x)
    Case lFILTER
        bind apply, SeqLib.filter(asFunc(x), seq(y))
    Case lCONCAT
        bind apply, SeqLib.concat(args)
    Case lLAST
        bind apply, SeqLib.last(seq(x))
    Case lBUTLAST = -33
        bind apply, SeqLib.butLast(seq(x))
    Case lNTH
        bind apply, SeqLib.nth(CLng(x), seq(y))
    Case lINC
        bind apply, x + 1
    Case lDEC
        bind apply, x - 1
    Case lconj
        bind apply, SeqLib.conj(seq(x), y)
    Case lRAND
        If nargs = 0 Then
            bind apply, Rnd()
        Else
            bind apply, Rnd() * x
        End If
    Case lRANDINT
        bind apply, CLng(Rnd() * x)
    Case lRANDBETWEEN
        bind apply, CLng(Rnd() * (y - x)) + x
    Case lREPEAT
        bind apply, SeqLib.repeat(x)
    Case lREPEATEDLY
        bind apply, SeqLib.repeatedly(asFunc(x))
    Case lLOADFILE
        bind apply, Lisp.readFile(CStr(x))
    Case lrange
        bind apply, SeqLib.seqRange(CLng(x))
'    Case lREVERSE
'        bind apply, SeqLib.seqReverse
    Case lsortby
        bind apply, SeqLib.seqSortBy(asFunc(x), seq(y))
    Case lsort
        If nil(y) Then
            bind apply, SeqLib.seqSort(seq(x))
        Else
            bind apply, SeqLib.seqSort(seq(x), truth(y))
        End If
    Case lconj
        bind apply, SeqLib.conj(seq(x), y)
End Select
End Function
Private Function numcompare(args As Collection, dir As String)
End Function
Private Function numcompute(args As Collection, op As String, initial As Single) As Single
Dim i As Long

numcompute = initial
Select Case op
    Case "X"
        For i = 1 To args.count
            If numcompute < args(i) Then numcompute = args(i)
        Next i
    Case "N"
        For i = 1 To args.count
            If numcompute > args(i) Then numcompute = args(i)
        Next i
    Case "+"
        For i = 1 To args.count
            numcompute = numcompute + args(i)
        Next i
    Case "-"
        numcompute = fst(args)
        For i = 2 To args.count
            numcompute = numcompute - args(i)
        Next i
    Case "*"
        For i = 1 To args.count
            numcompute = numcompute * args(i)
        Next i
    Case "/"
        numcompute = args(1)
        For i = 2 To args.count
            numcompute = numcompute / args(i)
        Next i
End Select
            
End Function

Private Function length(x As Variant) As Long
If isa(x, "Collection") Or isa(x, "Dictionary") Then
    length = x.count
ElseIf vartype(x) = vbString Then
    length = Len(x)
Else
    Err.Raise 101, , "Not implemented"
End If
End Function

'

'
'      ////////////////  SECTION 6.6 CHARACTERS
'    case CHARQ:           return truth(x instanceof Character);
'    case CHARALPHABETICQ: return truth(Character.isLetter(chr(x)));
'    case CHARNUMERICQ:    return truth(Character.isDigit(chr(x)));
'    case CHARWHITESPACEQ: return truth(Character.isWhitespace(chr(x)));
'    case CHARUPPERCASEQ:  return truth(Character.isUpperCase(chr(x)));
'    case CHARLOWERCASEQ:  return truth(Character.isLowerCase(chr(x)));
'    case CHARTOINTEGER:   return new Double((double)chr(x));
'    case INTEGERTOCHAR:   return chr((char)(int)num(x));
'    case CHARUPCASE:      return chr(Character.toUpperCase(chr(x)));
'    case CHARDOWNCASE:    return chr(Character.toLowerCase(chr(x)));
'    case CHARCMP+EQ:      return truth(charCompare(x, y, false) == 0);
'    case CHARCMP+LT:      return truth(charCompare(x, y, false) <  0);
'    case CHARCMP+GT:      return truth(charCompare(x, y, false) >  0);
'    case CHARCMP+GE:      return truth(charCompare(x, y, false) >= 0);
'    case CHARCMP+LE:      return truth(charCompare(x, y, false) <= 0);
'    case CHARCICMP+EQ:    return truth(charCompare(x, y, true)  == 0);
'    case CHARCICMP+LT:    return truth(charCompare(x, y, true)  <  0);
'    case CHARCICMP+GT:    return truth(charCompare(x, y, true)  >  0);
'    case CHARCICMP+GE:    return truth(charCompare(x, y, true)  >= 0);
'    case CHARCICMP+LE:    return truth(charCompare(x, y, true)  <= 0);
'
'    case ERROR:         return error(stringify(args));
'
'      ////////////////  SECTION 6.7 STRINGS
'    case STRINGQ:       return truth(x instanceof char[]);
'    case MAKESTRING:char[] str = new char[(int)num(x)];
'      if (y != null) {
'    char c = chr(y);
'    for (int i = str.length-1; i >= 0; i--) str[i] = c;
'      }
'      return str;
'    case STRING:        return listToString(args);
'    case STRINGLENGTH:  return num(str(x).length);
'    case STRINGREF:     return chr(str(x)[(int)num(y)]);
'    case STRINGSET:     Object z = third(args); str(x)[(int)num(y)] = chr(z);
'                        return z;
'    case SUBSTRING:     int start = (int)num(y), end = (int)num(third(args));
'                        return new String(str(x), start, end-start).toCharArray();
'    case STRINGAPPEND:  return stringAppend(args);
'    case STRINGTOLIST:  Pair result = null;
'                        char[] str2 = str(x);
'            for (int i = str2.length-1; i >= 0; i--)
'              result = cons(chr(str2[i]), result);
'            return result;
'    case LISTTOSTRING:  return listToString(x);
'    case STRINGCMP+EQ:  return truth(stringCompare(x, y, false) == 0);
'    case STRINGCMP+LT:  return truth(stringCompare(x, y, false) <  0);
'    case STRINGCMP+GT:  return truth(stringCompare(x, y, false) >  0);
'    case STRINGCMP+GE:  return truth(stringCompare(x, y, false) >= 0);
'    case STRINGCMP+LE:  return truth(stringCompare(x, y, false) <= 0);
'    case STRINGCICMP+EQ:return truth(stringCompare(x, y, true)  == 0);
'    case STRINGCICMP+LT:return truth(stringCompare(x, y, true)  <  0);
'    case STRINGCICMP+GT:return truth(stringCompare(x, y, true)  >  0);
'    case STRINGCICMP+GE:return truth(stringCompare(x, y, true)  >= 0);
'    case STRINGCICMP+LE:return truth(stringCompare(x, y, true)  <= 0);
'
'      ////////////////  SECTION 6.8 VECTORS
'    case VECTORQ:   return truth(x instanceof Object[]);
'    case MAKEVECTOR:    Object[] vec = new Object[(int)num(x)];
'                        if (y != null) {
'              for (int i = 0; i < vec.length; i++) vec[i] = y;
'            }
'            return vec;
'    case VECTOR:        return listToVector(args);
'    case VECTORLENGTH:  return num(vec(x).length);
'    case VECTORREF: return vec(x)[(int)num(y)];
'    case VECTORSET:     return vec(x)[(int)num(y)] = third(args);
'    case VECTORTOLIST:  return vectorToList(x);
'    case LISTTOVECTOR:  return listToVector(x);
'
'      ////////////////  SECTION 6.9 CONTROL FEATURES
'    case EVAL:          return interp.eval(x);
'    case FORCE:         return (!(x instanceof Procedure)) ? x
'              : proc(x).apply(interp, null);
'    case MACROEXPAND:   return Macro.macroExpand(interp, x);
'    case PROCEDUREQ:    return truth(x instanceof Procedure);
'    case APPLY:     return proc(x).apply(interp, listStar(rest(args)));
'    case MAP:           return map(proc(x), rest(args), interp, list(null));
'    case FOREACH:       return map(proc(x), rest(args), interp, null);
'    case CALLCC:        RuntimeException cc = new RuntimeException();
'                        Continuation proc = new Continuation(cc);
'                    try { return proc(x).apply(interp, list(proc)); }
'            catch (RuntimeException e) {
'                if (e == cc) return proc.value; else throw e;
'            }
'
'      ////////////////  SECTION 6.10 INPUT AND OUPUT
'    case EOFOBJECTQ:         return truth(x == InputPort.EOF);
'    case INPUTPORTQ:         return truth(x instanceof InputPort);
'    case CURRENTINPUTPORT:   return interp.input;
'    case OPENINPUTFILE:      return openInputFile(x);
'    case CLOSEINPUTPORT:     return inPort(x, interp).close();
'    case OUTPUTPORTQ:        return truth(x instanceof PrintWriter);
'    case CURRENTOUTPUTPORT:  return interp.output;
'    case OPENOUTPUTFILE:     return openOutputFile(x);
'    case CALLWITHOUTPUTFILE: PrintWriter p = null;
'                             try { p = openOutputFile(x);
'                                   z = proc(y).apply(interp, list(p));
'                             } finally { if (p != null) p.close(); }
'                             return z;
'    case CALLWITHINPUTFILE:  InputPort p2 = null;
'                             try { p2 = openInputFile(x);
'                                   z = proc(y).apply(interp, list(p2));
'                             } finally { if (p2 != null) p2.close(); }
'                             return z;
'    case CLOSEOUTPUTPORT:    outPort(x, interp).close(); return TRUE;
'    case READCHAR:      return inPort(x, interp).readChar();
'    case PEEKCHAR:      return inPort(x, interp).peekChar();
'    case LOAD:          return interp.load(x);
'    case READ:      return inPort(x, interp).read();
'    case EOF_OBJECT:    return truth(InputPort.isEOF(x));
'    case WRITE:     return write(x, outPort(y, interp), true);
'    case DISPLAY:       return write(x, outPort(y, interp), false);
'    case NEWLINE:   outPort(x, interp).println();
'                        outPort(x, interp).flush(); return TRUE;
'
'      ////////////////  EXTENSIONS
'    case CLASS:         try { return Class.forName(stringify(x, false)); }
'                        catch (ClassNotFoundException e) { return FALSE; }
'    case NEW:           try { return JavaMethod.toClass(x).newInstance(); }
'                        catch (ClassCastException e)     { ; }
'                        catch (NoSuchMethodError e)      { ; }
'                        catch (InstantiationException e) { ; }
'                        catch (ClassNotFoundException e) { ; }
'                        catch (IllegalAccessException e) { ; }
'                        return FALSE;
'    case METHOD:        return new JavaMethod(stringify(x, false), y,
'                          rest(rest(args)));
'    case EXIT:          System.exit((x == null) ? 0 : (int)num(x));
'    case LISTSTAR:      return listStar(args);
'    case TIMECALL:      Runtime runtime = Runtime.getRuntime();
'                        runtime.gc();
'                        long startTime = System.currentTimeMillis();
'            long startMem = runtime.freeMemory();
'            Object ans = FALSE;
'            int nTimes = (y == null ? 1 : (int)num(y));
'            for (int i = 0; i < nTimes; i++) {
'              ans = proc(x).apply(interp, null);
'            }
'                        long time = System.currentTimeMillis() - startTime;
'            long mem = startMem - runtime.freeMemory();
'            return cons(ans, list(list(num(time), "msec"),
'                          list(num(mem), "bytes")));
'    default:            return error("internal error: unknown primitive: "
'                     + this + " applied to " + args);
'    }
'    }
'
'  public static char[] stringAppend(Object args) {
'    StringBuffer result = new StringBuffer();
'    for(; args instanceof Pair; args = rest(args)) {
'      result.append(stringify(first(args), false));
'    }
'    return result.toString().toCharArray();
'  }
'
'  public static Object memberAssoc(Object obj, Object list, char m, char eq) {
'    while (list instanceof Pair) {
'      Object target = (m == 'm') ? first(list) : first(first(list));
'      boolean found;
'      switch (eq) {
'      case 'q': found = (target == obj); break;
'      case 'v': found = eqv(target, obj); break;
'      case ' ': found = equal(target, obj); break;
'      default: warn("Bad option to memberAssoc:" + eq); return FALSE;
'      }
'      if (found) return (m == 'm') ? list : first(list);
'      list = rest(list);
'    }
'    return FALSE;
'  }
'
'  public static Object numCompare(Object args, char op) {
'    while (rest(args) instanceof Pair) {
'      double x = num(first(args)); args = rest(args);
'      double y = num(first(args));
'      switch (op) {
'      case '>': if (!(x >  y)) return FALSE; break;
'      case '<': if (!(x <  y)) return FALSE; break;
'      case '=': if (!(x == y)) return FALSE; break;
'      case 'L': if (!(x <= y)) return FALSE; break;
'      case 'G': if (!(x >= y)) return FALSE; break;
'      default: error("internal error: unrecognized op: " + op); break;
'      }
'    }
'    return TRUE;
'  }
'
'  public static Object numCompute(Object args, char op, double result) {
'    if (args == null) {
'      switch (op) {
'      case '-': return num(0 - result);
'      case '/': return num(1 / result);
'      default:  return num(result);
'      }
'    } else {
'      while (args instanceof Pair) {
'    double x = num(first(args)); args = rest(args);
'    switch (op) {
'    case 'X': if (x > result) result = x; break;
'    case 'N': if (x < result) result = x; break;
'    case '+': result += x; break;
'    case '-': result -= x; break;
'    case '*': result *= x; break;
'    case '/': result /= x; break;
'    default: error("internal error: unrecognized op: " + op); break;
'    }
'      }
'      return num(result);
'    }
'  }
'
'  /** Return the sign of the argument: +1, -1, or 0. **/
'  static int sign(int x) { return (x > 0) ? +1 : (x < 0) ? -1 : 0; }
'
'  /** Return <0 if x is alphabetically first, >0 if y is first,
'   * 0 if same.  Case insensitive iff ci is true.  Error if not both chars. **/
'  public static int charCompare(Object x, Object y, boolean ci) {
'    char xc = chr(x), yc = chr(y);
'    if (ci) { xc = Character.toLowerCase(xc); yc = Character.toLowerCase(yc); }
'    return xc - yc;
'  }
'
'  /** Return <0 if x is alphabetically first, >0 if y is first,
'   * 0 if same.  Case insensitive iff ci is true.  Error if not strings. **/
'  public static int stringCompare(Object x, Object y, boolean ci) {
'    if (x instanceof char[] && y instanceof char[]) {
'      char[] xc = (char[])x, yc = (char[])y;
'      for (int i = 0; i < xc.length; i++) {
'    int diff = (!ci) ? xc[i] - yc[i]
'      : Character.toUpperCase(xc[i]) - Character.toUpperCase(yc[i]);
'    if (diff != 0) return diff;
'      }
'      return xc.length - yc.length;
'    } else {
'      error("expected two strings, got: " + stringify(list(x, y)));
'      return 0;
'    }
'  }
'
'  static Object numberToString(Object x, Object y) {
'    int base = (y instanceof Number) ? (int)num(y) : 10;
'    if (base != 10 || num(x) == Math.round(num(x))) {
'      // An integer
'      return Long.toString((long)num(x), base).toCharArray();
'    } else {
'      // A floating point number
'      return x.toString().toCharArray();
'    }
'  }
'
'  static Object stringToNumber(Object x, Object y) {
'    int base = (y instanceof Number) ? (int)num(y) : 10;
'    try {
'      return (base == 10)
'    ? Double.valueOf(stringify(x, false))
'    : num(Long.parseLong(stringify(x, false), base));
'    } catch (NumberFormatException e) { return FALSE; }
'  }
'
'  static Object gcd(Object args) {
'    long gcd = 0;
'    while (args instanceof Pair) {
'      gcd = gcd2(Math.abs((long)num(first(args))), gcd);
'      args = rest(args);
'    }
'    return num(gcd);
'  }
'
'  static long gcd2(long a, long b) {
'    if (b == 0) return a;
'    else return gcd2(b, a % b);
'  }
'
'  static Object lcm(Object args) {
'    long L = 1, g = 1;
'    while (args instanceof Pair) {
'      long n = Math.abs((long)num(first(args)));
'      g = gcd2(n, L);
'      L = (g == 0) ? g : (n / g) * L;
'      args = rest(args);
'    }
'    return num(L);
'  }
'
'  static boolean isExact(Object x) {
'    if (!(x instanceof Double)) return false;
'    double d = num(x);
'    return (d == Math.round(d) && Math.abs(d) < 102962884861573423.0);
'  }
'
'  static PrintWriter openOutputFile(Object filename) {
'    try {
'      return new PrintWriter(new FileWriter(stringify(filename, false)));
'    } catch (FileNotFoundException e) {
'      return (PrintWriter)error("No such file: " + stringify(filename));
'    } catch (IOException e) {
'      return (PrintWriter)error("IOException: " + e);
'    }
'  }
'
'  static InputPort openInputFile(Object filename) {
'    try {
'      return new InputPort(new FileInputStream(stringify(filename, false)));
'    } catch (FileNotFoundException e) {
'      return (InputPort)error("No such file: " + stringify(filename));
'    } catch (IOException e) {
'      return (InputPort)error("IOException: " + e);
'    }
'  }
'
'  static boolean isList(Object x) {
'    Object slow = x, fast = x;
'    for(;;) {
'      if (fast == null) return true;
'      if (slow == rest(fast) || !(fast instanceof Pair)
'      || !(slow instanceof Pair)) return false;
'      slow = rest(slow);
'      fast = rest(fast);
'      if (fast == null) return true;
'      if (!(fast instanceof Pair)) return false;
'      fast = rest(fast);
'    }
'  }
'
'  static Object append(Object args) {
'    if (rest(args) == null) return first(args);
'    else return append2(first(args), append(rest(args)));
'  }
'
'  static Object append2(Object x, Object y) {
'    if (x instanceof Pair) return cons(first(x), append2(rest(x), y));
'    else return y;
'  }
'
'  /** Map proc over a list of lists of args, in the given interpreter.
'   * If result is non-null, accumulate the results of each call there
'   * and return that at the end.  Otherwise, just return null. **/
'  static Pair map(Procedure proc, Object args, Scheme interp, Pair result) {
'    Pair accum = result;
'    if (rest(args) == null) {
'      args = first(args);
'      while (args instanceof Pair) {
'    Object x = proc.apply(interp, list(first(args)));
'    if (accum != null) accum = (Pair) (accum.rest = list(x));
'    args = rest(args);
'      }
'    } else {
'      Procedure car = proc(interp.eval("car")), cdr = proc(interp.eval("cdr"));
'      while  (first(args) instanceof Pair) {
'    Object x = proc.apply(interp, map(car, list(args), interp, list(null)));
'    if (accum != null) accum = (Pair) (accum.rest = list(x));
'    args = map(cdr, list(args), interp, list(null));
'      }
'    }
'    return (Pair)rest(result);
'  }
'
'}


'/** Map proc over a list of lists of args, in the given interpreter.
' * If result is non-null, accumulate the results of each call there
' * and return that at the end.  Otherwise, just return null. **/



Private Function IFn_apply(args As Collection) As Variant
bind IFn_apply, apply(args)
End Function