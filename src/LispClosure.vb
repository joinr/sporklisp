Option Explicit

'
'public class Closure extends Procedure {
'
'    Object parms;
'    Object body;
'    Environment env;
'
'    /** Make a closure from a parameter list, body, and environment. **/
'    public Closure (Object parms, Object body, Environment env) {
'        this.parms = parms;
'    this.env = env;
'    this.body = (body instanceof Pair && rest(body) == null)
'        Print first(body)
'        : cons("begin", body);
'    }
'
'    /** Apply a closure to a list of arguments.  **/
'    public Object apply(Scheme interpreter, Object args) {
'    return interpreter.eval(body, new Environment(parms, args, env));
'    }
'}

Public name As String

Public params As Collection
Public body As Variant
Public env As Dictionary
Implements IFn

Public Function apply(args As Collection) As Variant
bind apply, Lisp.eval(body, makeEnvironment(params, args, env))
End Function

Private Sub Class_Initialize()
name = "Anonymous Function"
End Sub


Private Function IFn_apply(args As Collection) As Variant
bind IFn_apply, apply(args)
End Function