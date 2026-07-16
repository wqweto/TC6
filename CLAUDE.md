# TC6 (TwinClient 6) — SQLite replacement

## Coding Style for this project

- Declare all helper procedures `Private`
- Use `On Error GoTo EH` where needed; in the handler `Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"`
- Use `Long` for array indices, sizes and counters (avoid `Integer`)
- Avoid `Variant`; avoid dynamic arrays in inner loops (pre-allocate)
- One blank line between procedures
- One blank line between the `Dim` block and the first executable line inside a procedure; no other blank lines inside a procedure body
- Comments only where the VB6 diverges non-obviously from the C (e.g. unsigned workarounds)
- Use hungarian notation: `s` - String, `l` - Long, `n` - Integer, `b` - Boolean, `o` - Object, `c` - Collection, `d` - Date, `dbl` - Double, `sng` - Single, `byt` - Byte, `u` - UDTs, `h` - Handles (incl. hResult), `cy` - Currency, `e` - Enums, `p` - interface pointers
- Use `ba` prefix for byte arrays and `a` for all other arrays regardless of element type
- Use `m_` prefix for member variables and `g_` for global ones
- Use `md` prefix for standard modules, `c` for classes, `frm` for forms, `ctx` for user-controls
- Declare all variables at the beginning of the procedure, separated from the code by a blank line
- One local variable declaration per line, data-type aligned at column 25
- Align API consts data-types at column 45
- Align module variables data-types at column 37
- Declare API consts local to a routine if not used in any other routine
- API declares use "dllname" without the .dll suffix, always Unicode versions (aliased to names without the W suffix)
- Separate logical sections inside a procedure with `'---` comments, not blank lines
- Start comments with `'---` instead of a single `'`, except for a comment banner at the start of a procedure/module
- Put only one statement per logical line i.e. don't use : to separate multiple statements
- Put `If` statements on separate lines i.e. don't merge `If Cond Then Stmt` on a single line
- Use `Call` only for API functions whose result is discarded (not used in an `If`); never otherwise
- Use `QH` label for cleanup before `EH` label
- Align `Case`es after `Select Case` at the same column i.e. don't indent
- All `Long` const hex literals between &H8000 and &HFFFF must use the & type character (e.g. &H8000&) or risk being sign-extended
- Never use `Next Var` i.e. just `Next`
- Always omit `ByRef`; specify `ByVal` only when needed
- Order of procedures in module: public events/enums/types, API declares, member variables and private enums/types, properties, methods, event handlers, base class events (e.g. Class_Terminate)
- Within properties/methods order by visibility: public, friend, private
- Use `pv` prefix for private procedures and `fr` for friend ones
- Use `lIdx`, `lJdx`, etc. instead of single-letter index variable `i`, `j`, etc.
- ReDim uses explicit data-type i.e. `ReDim aName(0 To 100) As String`
