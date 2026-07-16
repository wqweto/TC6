# TC6 (TwinClient 6) — Working Notes

## Goal

Reimplement the SQLite subsystem of Olaf Schmidt's **vbRichClient6 (RC6)** COM
library as a native **TwinBasic** project. The RC6 type library is provided in
IDL form at [doc/RC6.idl](doc/RC6.idl) (~15,700 lines, the whole RC6 surface).

This first pass extracts the SQLite-related coclasses from the typelib and lays
them down as VB6 `.cls` stubs under [src/](src/) — full method/property/event
signatures, empty bodies — so the implementation can be filled in against a
fixed, binary-compatible interface.

## Prompts (chronological)

1. **Initial task** — *"This is a RC6 sqlite replacement project for TwinBasic.
   Analyze RC6 typelib and get list of classes which have to do with sqlite and
   stub all of these with methods/properties/events in src folder as VB6 .cls
   files first."*
2. **Follow-up** — *"Document all the work done and prompts in NOTES.md"* (this
   file).
3. **API module** — *"Create shared mdWinApi.bas with Public Declares/Consts and
   include all exports of built-in winsqlite3.dll. For future const reference use
   sqlite3.h in doc (don't dump all consts in mdWinApi now)."*
4. **Correction** — *"I think the exports are stdcall"* → verified true (see the
   Calling convention section below); declares regenerated as StdCall.
5. **VB6 compat** — *"Make code VB6 compatible in x86 version i.e. LongLong to
   Currency or split to two Longs"* and *"Don't use CDecl APIs in both x86 and
   x64. Tell me if this is not possible"* → done: `LongLong`→`Currency`, the 8
   CDecl/variadic exports dropped (possible — none are on the data-access path).
6. **LongPtr consts** — *"For LongPtr consts just declare LongPtr in x64 and Long
   in x86"* → `SQLITE_STATIC`/`SQLITE_TRANSIENT` wrapped in `#If Win64`.
7. **Test project** — *"Create a test Std-EXE VB6 project and try to compile
   sources"* → [test/apitest.vbp](test/apitest.vbp); compiles and runs (below).

## How the SQLite classes were identified

RC6 is a large general-purpose library (Cairo graphics, TCP/UDP, audio, physics,
WebView2, etc.). The SQLite subsystem was isolated by:

- Searching the IDL for `sqlite` (hits: `GetSqliteErrStr`, the `FieldType`
  enum with `SQLite_INTEGER…SQLite_NULL`, `CreateTableFromRs`).
- Following the object graph outward from `cConnection` — the SQLite DB handle
  wrapper — through everything it returns or consumes (recordsets, commands,
  schema objects, UDF interfaces).

## Classes stubbed (26)

### Core data access
| Class | Instancing | Notes |
|-------|-----------|-------|
| `cConnection` | Creatable | Main SQLite DB wrapper. Open/Create/Copy DB, transactions (with savepoints), commands, UDF/collation registration, encryption (codec). Raises transaction events. |
| `cRecordset` | Creatable | Disconnected, updatable recordset. Navigation, find, sort, batch update, JSON/ADO export. Raises change events. |
| `cCommand` | PublicNotCreatable | Parameterised write/DML statement (bind `Set*`, `Execute`). |
| `cSelectCommand` | PublicNotCreatable | Parameterised read statement returning a `cRecordset`. |
| `cCursor` | PublicNotCreatable | Forward-only low-level cursor (`Step`/`Reset`, column accessors). |

### Schema objects (all PublicNotCreatable)
| Class | Notes |
|-------|-------|
| `cDataBase` / `cDataBases` | Attached databases; `Tables`, `Views`. |
| `cTable` / `cTables` | Table schema; `Columns`, `Indexes`, `Triggers`. |
| `cColumn` / `cColumns` | Column definition (constraints, conflict algorithms). |
| `cIndex` / `cIndexes` | Index name + DDL. |
| `cTrigger` / `cTriggers` | Trigger name + DDL. |
| `cView` / `cViews` | View name + DDL. |
| `cField` / `cFields` | Runtime recordset fields (value + metadata). |

### Helpers / UDF
| Class | Instancing | Notes |
|-------|-----------|-------|
| `cMemDB` | Creatable | Convenience in-memory DB (aggregate helpers `GetSum`/`GetAvg`/…). |
| `cDBAccess` | Creatable | Thread-marshalled DB access (RC6 threading). |
| `cConverter` | Creatable | ADO → SQLite migration; raises progress events. |
| `cUDFMethods` | PublicNotCreatable | Passed into UDF callbacks to read args / set result. |
| `IFunction` | Creatable | Implement to register a scalar user-defined function. |
| `IAggregateFunction` | Creatable | Implement for aggregate UDF (`Step`/`Final`). |
| `ICollation` | Creatable | Implement for a custom collation sequence. |

## Public enums (declared inside the class that primarily uses them)

Placed in `.cls` files rather than a separate module, per the "*.cls files
first*" instruction. In VB6/TwinBasic a `Public Enum` in a public class is
visible project-wide.

| Enum | Values | Declared in |
|------|--------|-------------|
| `eConnectionSynchronousType` | `SynchronousOff/Normal/Full/Extra` | `cConnection` |
| `eCodecType` | `CODEC_TYPE_AES128/AES256/CHACHA20/SQLCIPHER/RC4` | `cConnection` |
| `FieldType` | `SQLite_INTEGER…SQLite_NULL`, `VB_*_AutoConverted` | `cField` |
| `ConflictAlgorithms` | `OnConflict_UseTableDefault…OnConflict_REPLACE` | `cColumn` |
| `ResetEnum` | `ResetAll/DeletesOnly/InsertsOnly/UpdatesOnly` | `cRecordset` |

## Events

Sourced from the outgoing (`[source]`) dispinterfaces in the IDL:

- `cConnection`: `CommitTransComplete`, `RollbackTransComplete`,
  `CommitForSavePointComplete`, `SavePointReleased`,
  `RollbackToSavePointComplete`.
- `cRecordset`: `FieldChange`, `Move`, `AddNew`, `Delete`, `QueryFinished`,
  `Reset`.
- `cConverter`: `SchemaProgress`, `InsertProgress`, `IndexProgress`.

### Event timing (probed against RC6 6.0.15, pinned by mdEventTests)

Captured by driving identical scenarios against both engines with a
VBScript event sink (probe scripts in the session scratchpad); TC6 matches
RC6 trace-for-trace:

- `Move`/`AddNew`/`Delete` args are **0-based row indices**; BOF = **-1**,
  EOF = **-2**. Move-family events fire **only when the position actually
  changes** (`MoveFirst` when already on row 0 is silent).
- `MoveNext` at EOF / `MovePrevious` at BOF raise `vbObjectError`
  (&H80040000) with no event; `MoveFirst`/`MoveLast` on an empty recordset
  are silent no-ops.
- `AbsolutePosition` is **1-based**; Get: -1 = empty rs, -2 = BOF, -3 =
  EOF; Let: out of range (incl. 0) raises `vbObjectError`, same-row
  assignment is silent, otherwise a single 0-based `Move`.
- `Delete` carries the **new** position: mid-delete keeps the index (next
  row slides in), last-row delete clamps to the new last row, deleting the
  only row reports -1 with BOF+EOF set.
- `AddNew` raises `AddNew(newIdx)` only — no `Move` despite the cursor
  landing on the new row. Field sets then fire `FieldChange(row, col)`.
- Assigning a Field its **current value is a complete no-op** — no
  `FieldChange`, no dirty flag (RC6 compares before marking).
- `Sort` Let rewinds to the first row and always raises a single
  `Move(0)` (even when already on row 0, and also when clearing the sort).
  `FindFirst`/etc raise `Move(idx)` on a hit, nothing on a miss.
- `ReQuery` raises `QueryFinished` only; the cursor rewinds to the first
  row without a `Move`. `Content` Let raises nothing. `ResetChanges`
  raises `Reset`. `UpdateBatch` raises no recordset events.

**Documented divergences** (deliberate, asserted as TC6 behavior in
mdEventTests):

- **Named savepoints**: RC6 6.0.15's `BeginTrans`/`CommitTrans`/
  `RollbackTrans` savepoint names are dead code — probed with data checks:
  nesting is a plain counter, a named `RollbackTrans` rolls back the whole
  transaction (`RollbackTransComplete`), an inner named `CommitTrans` is a
  silent counter decrement, and `CommitForSavePointComplete`/
  `SavePointReleased`/`RollbackToSavePointComplete` never fire. TC6
  implements the real SAVEPOINT semantics the interface declares (and
  raises those events); unnamed transactions match RC6 exactly. NB: both
  engines silently no-op `CommitTrans`/`RollbackTrans` with no open txn.
- **`SortRefresh`**: RC6's cursor drifts erratically (+1 row on a clean
  recordset, no re-sort at all with pending dirty edits) — evidently
  buggy; TC6 re-applies the sort with the cursor pinned to its record,
  raising `Move` only if the index changed.
- **`ResetChanges` under an active sort**: RC6 emits a Move storm from its
  row-chain rebuild and re-sorts on original values; TC6 re-applies the
  sort (record-pinned, at most one `Move`) and always raises `Reset`.
- **`Sort = ""`**: RC6 restores the natural row order (nondestructive row
  chain); TC6's in-place permute keeps the current order (still rewinds
  with `Move(0)`).

### Error contract (probed against RC6 6.0.15, pinned by mdErrorTests)

37 scenarios trigger every user-facing `Err.Raise` identically on TC6 and
the live RC6.dll and compare `Err.Number & "|" & Err.Description`. RC6
never populates `Err.Source` (VB stamps the project name), so Source is
excluded from the comparison; TC6 fills it with `Class.Method` anyway.

- **SQL engine errors** carry a prefix + the raw sqlite message:
  `Cannot compile Select-Statement: ` (OpenRecordset/select prepare),
  `Cannot compile SQL-Statement: ` (Execute/ExecCmd/cCommand prepare),
  `Cannot execute SQL-Statement: ` (any DML step failure). A step failure
  while *materialising a select* additionally appends the result code:
  `Cannot execute Select-Statement: integer overflow (1)` — only there.
  `cConnection.Execute` prepares/steps each statement itself (not
  `sqlite3_exec`) so the two prefixes can be told apart.
- **Recordset state errors** (all `vbObjectError`) replicate RC6's exact
  strings *including two RC6 typos*: `The OneBased Index is out of range`,
  `The Recordset has already reached EOF`/`BOF`, `The Recordset is
  postioned at EOF`/`BOF` (Delete, sic), `Either BOF of EOF is True`
  (cell access, sic), `Recordset is not updatable`.
- **Silent ignores**: RC6 raises *nothing* when assigning `Value` on a
  non-updatable recordset or an expression column — the write is simply
  dropped (TC6: early `Exit Sub` in `frSetCellValueAt`).
- **Field lookups** (`cFields`, also `Sort`): unknown name raises
  `vbObjectError` `No such Field-Def: <name>` (an unclosed `[` sort token
  is treated as a literal field name, e.g. `No such Field-Def: [unclosed`),
  index out of range raises `Field-Index out of range`. Bad find criteria:
  `This is not a valid Find-Criterion`.
- **Schema collections** (`cTables`/`cColumns`/`cIndexes`/`cViews`/
  `cTriggers`) raise a bare error **9** (`Subscript out of range`) on a
  name miss; `cDataBases` raises bare **5**. VB6 trap: `Err.Raise 9`
  inside an active error handler *reuses the pending `Err.Description`*
  (the VBA.Collection miss text) — `Err.Clear` first or the default
  description is wrong.
- **Commands**: `cSelectCommand` demands at least one parameter
  (`?`/`:n`/`@n`/`$n`) — otherwise `Couldn't find any Parameters in the
  query-string` — but accepts an *invalid* statement silently, deferring
  the compile (and its `Cannot compile Select-Statement` error) to
  `Execute`. VB6 trap there: an error swallowed by `On Error Resume Next`
  lingers in the global `Err` object after the Sub exits — follow with
  `On Error GoTo 0`. `cCommand` compiles eagerly. Bind errors diverge by
  class: `cCommand` raises `Cannot set Parameter: <msg>`, `cSelectCommand`
  raises a bare error 9.

## IDL → VB6 type mapping used

| IDL | VB6 |
|-----|-----|
| `BSTR` | `String` |
| `long` / `double` / `DATE` / `short` | `Long` / `Double` / `Date` / `Integer` |
| `VARIANT` / `VARIANT_BOOL` | `Variant` / `Boolean` |
| `unsigned char` | `Byte` |
| `SAFEARRAY(unsigned char / VARIANT / BSTR)` | `Byte()` / `Variant()` / `String()` |
| `_cXxx**` (RC6 interface) | corresponding `cXxx` class |
| `IUnknown**` | `IUnknown` (the `NewEnum` collection members) |
| `IDispatch**` / `_Connection**` (ADO) | `Object` (avoids an external ADO reference) |
| `_Collection` (VBA) | `Collection` |
| `[in]` scalar | `ByVal` |
| `[in, out]` | `ByRef` (default, so `ByVal` omitted) |
| `[out, retval]` | function return value |
| `vararg` | `ParamArray … As Variant` |

## Conventions applied (per CLAUDE.md)

- Standard VB6 `.cls` header block; `VB_Creatable = True` for creatable classes,
  `False` for the `noncreatable` coclasses.
- Hungarian parameter names kept as-is from the IDL for signature fidelity.
- `ByRef` omitted; `ByVal` only where the IDL marks `[in]` scalars.
- Default members carry `Attribute …VB_UserMemId = 0` (collection `Item`,
  `cField.Value`, `cRecordset.Fields`); `NewEnum` carries `VB_UserMemId = -4`
  so `For Each` works.

## [src/sqlite3win32stubs.bas](src/sqlite3win32stubs.bas) — winsqlite3.dll API layer

(originally `mdWinApi.bas`, then `mdSqliteApi.bas`; renamed once it held only the `stub_sqlite3_*` declares —
the UTF-8 helpers moved to [src/mdGlobals.bas](src/mdGlobals.bas).)

Shared standard module holding the raw SQLite API surface. Generated by dumping
the export table of the OS-supplied `C:\Windows\System32\winsqlite3.dll`
(SQLite 3.x, **297** exports) and parsing the prototypes from
[doc/sqlite3.h](doc/sqlite3.h).

- **283** functions declared (`Public Declare … Lib "winsqlite3"`), later reduced
  to 275 (StdCall-only). Each is named `stub_sqlite3_Xxx` with `Alias
  "sqlite3_Xxx"` (originally `vbsqlite3_Xxx`, renamed) — the `stub_` prefix
  avoids clashing with any real `sqlite3_*` symbol and keeps
  call sites greppable, while the Alias binds to the true DLL export.
- **14** exports are data symbols or internal helpers with no public prototype
  (`sqlite3_version[]`, `sqlite3_temp_directory`, `sqlite3_data_directory`,
  `sqlite3_fts3_may_be_corrupt`, `sqlite3_unsupported_selecttrace`, and the
  internal `sqlite3_win32_*` helpers) — listed in a footer comment, declare when
  needed.
- Only a **curated core set of constants** is included (result codes, datatypes,
  text encodings, open/prepare flags, `SQLITE_STATIC`/`SQLITE_TRANSIENT`),
  split out into [src/sqlite3win32helper.bas](src/sqlite3win32helper.bas) —
  the stubs module holds declares only (plus the `LongPtr` shim they need).
  The full constant set stays in `doc\sqlite3.h` per the instruction.

### VB6 (x86) / twinBASIC dual-target compatibility

The module compiles unchanged in VB6 (x86) and tB (x86/x64):

- **No `PtrSafe`** — VB6 rejects it ("Expected: Sub or Function"); tB does not
  require it. Omitted from every declare.
- **`LongPtr`** — defined as an enum shim for VB6, skipped under tB/x64:
  `#If Win64 = 0 And TWINBASIC = 0 Then Public Enum LongPtr : [_] : End Enum`.
  On 32-bit VB6 that is a 4-byte Long = pointer size.
- **`sqlite3_int64`** — the 26 int64 declares are wrapped `#If Win64 Then …
  LongLong #Else … Currency #End If`: `LongLong` on x64 (tB), `Currency` on x86
  (VB6, which lacks `LongLong`: "User-defined type not defined"). Both are 8-byte;
  on x86 put the raw 64-bit value into the `Currency` via a VarPtr/CopyMemory
  helper (wrapper concern).
- **UTF-8 marshaling** lives in `mdGlobals.bas`: `ToUtf8Array` (String → UTF-8
  `Byte()`) and `FromUtf8Array` (UTF-8 `Byte()` → String), over
  `WideCharToMultiByte`/`MultiByteToWideChar` with `CP_UTF8`. `sqlite3win32stubs.bas`
  is kept declares-only.
- **`SQLITE_STATIC` / `SQLITE_TRANSIENT`** — `#If Win64 Then As LongPtr #Else As
  Long`. VB6 will not type a `Const` as an enum, and `-1` sign-extends correctly.
- **The 8 CDecl/variadic exports are not declared at all** — keeps the module
  pure StdCall so nothing depends on mixed calling conventions.

### Calling convention (verified, not assumed)

winsqlite3.dll is built **StdCall**, which is why it is callable from a plain
VB6/tB `Declare` at all. Confirmed by disassembly (32-bit `SysWOW64` build):
non-variadic functions clean their own stack — e.g. `sqlite3_busy_timeout`
(2 args) ends `ret 8`.

The **exception** is the printf-style C variadic (`...`) family, which is
necessarily **CDecl** — confirmed: `sqlite3_mprintf`, `sqlite3_snprintf`,
`sqlite3_config`, `sqlite3_db_config`, `sqlite3_log`, `sqlite3_test_control`,
`sqlite3_vtab_config`, `sqlite3_str_appendf` all end in a plain `ret`. These
eight carry the `CDecl` keyword and are declared with their fixed arguments only.
The `va_list` variants (`sqlite3_vmprintf`, `sqlite3_vsnprintf`,
`sqlite3_str_vappendf`) are **StdCall** (fixed arity — e.g. `vmprintf` ends
`ret 8`).

### C → VB type mapping (declares)

| C | VB / tB |
|---|---------|
| `int`, `unsigned` | `Long` |
| `sqlite3_int64` / `sqlite3_uint64` | `LongLong` |
| `double` | `Double` |
| any pointer / handle / `const char*` / callback / `va_list` | `LongPtr` (all `ByVal`) |
| `void` return | `Sub` |

Strings are **UTF-8**: callers pass a `LongPtr` to a UTF-8 byte buffer, never a
VB `String` (which is UTF-16).

## [test/TestRunner.vbp](test/TestRunner.vbp) — ActiveX-DLL test runner

An **ActiveX DLL** (`Type=OleDll`, project name `TC6SQLiteTest`) that compiles
**all 26 classes + both shared modules** plus the test code, and exposes one
public COM class, `cTestHost`, to drive the suite from PowerShell/VBScript. This
is the single project going forward — the old `apitest` Std-EXE was removed (its
API smoke-check is now covered by `cConnection.CreateAndInsert`). Current run:
**8 tests / 33 checks, all PASS** against winsqlite3.dll.

- [test/cTestHost.cls](test/cTestHost.cls) — public `MultiUse` class, the COM
  entry point. `RunAll([OutputFile])` and `RunTests([Filter],[OutputFile])`
  return the failed-test count; `Report`/`TestsRun`/`TestsFailed` properties
  expose results. `Filter` = comma-separated, case-insensitive substrings of
  test names; empty = all. `OutputFile` empty → `<dll folder>\testrun.log`.
- [test/mdTestRunner.bas](test/mdTestRunner.bas) — the engine.
  `TestReset`/`TestBegin`/`TestEnd`/`TestErr`/`TestFinish` +
  `AssertTrue`/`AssertEqLng`/`AssertEqStr` + report buffer. `TestBegin(name)`
  returns `False` when the name is filtered out, so each `Test_` sub starts with
  `If Not TestBegin(name) Then Exit Sub` and traps its own errors — one
  failing/skipped test never aborts the rest.
- [test/mdConnectionTests.bas](test/mdConnectionTests.bas) — one standard module
  **per class** holds that class's `Test_*` subs + a public `RunXxxTests` entry
  that `cTestHost.RunTests` calls. Add a new module + one call for each class.
- [test/run_tests.vbs](test/run_tests.vbs) — convenience driver (see below).
- [test/TestRunnerBin.vbp](test/TestRunnerBin.vbp) — the **same test modules
  compiled against the registered release `TC6SQLite.dll`** (typelib
  reference instead of source classes), project name `TC6SQLiteTestBin`;
  `run_tests.vbs /bin` drives it. Exercises the statically linked SQLite in
  the shipped binary rather than winsqlite3. The only shims needed
  ([test/mdBinShims.bas](test/mdBinShims.bas)): `FromUtf8Array` and the
  `Public Enum LongPtr` x86 trick (both live in bas modules, which don't
  make it into the typelib) — the tests otherwise use only public API.

**Build (headless):** `Start-Process VB6.EXE -ArgumentList '/make',<vbp>,'/out',
<log> -Wait` — VB6 is GUI-subsystem so the shell must wait; the `/out` log is
append-mode. Building an ActiveX DLL **auto-registers** it on this machine
(No/Project compatibility, `CompatibleMode=0`, so CLSIDs churn each build — the
stable `TC6SQLiteTest.cTestHost` ProgID always resolves to the latest build).

**Run (the DLL is x86 → use a 32-bit script host):**

```
C:\Windows\SysWOW64\cscript.exe //nologo test\run_tests.vbs [/out:log] [/filter:tokens]
```

`run_tests.vbs` does `CreateObject("TC6SQLiteTest.cTestHost")`, calls
`RunTests(filter, out)`, echoes `.Report`, and `WScript.Quit`s the failed count.
Equivalently from 32-bit PowerShell:
`& $env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -Command "(New-Object -ComObject TC6SQLiteTest.cTestHost).RunAll('out.log')"`.
NB: invoke `cscript` with the call operator, not `Start-Process` — the latter
mangled the `/out:C:\…` argument.

**Two project requirements** beyond a bare vbp: the
`Reference=…stdole2.tlb#OLE Automation` line (the classes use `IUnknown` for the
`NewEnum` members — without it VB6 reports *"User-defined type not defined"*),
and correct **class instancing** (restored below).

### Class instancing (restored for the ActiveX DLL)

The 8 creatable classes (`cConnection`, `cRecordset`, `cMemDB`, `cConverter`,
`cDBAccess`, `IFunction`, `IAggregateFunction`, `ICollation`) are `MultiUse`
(`MultiUse = -1 / VB_Creatable = True / VB_Exposed = True`); the other 18 are
`PublicNotCreatable` (`MultiUse = 0 / VB_Creatable = False / VB_Exposed = True`).
This also fixed a latent stub bug: the schema/child classes had an **invalid**
header (`MultiUse = -1` with `VB_Creatable = False`) that loaded in no project
type. `cTestHost` is `MultiUse`.

## Status / next steps

- [x] SQLite classes identified from the typelib.
- [x] All 26 classes stubbed with full signatures, empty bodies.
- [x] SQLite backend binding: `sqlite3win32stubs.bas` with all winsqlite3.dll exports
      declared (StdCall, verified) plus core constants; UTF-8 helpers in
      `mdGlobals.bas`.
- [x] VB6 (x86) + tB dual-target compatibility of `sqlite3win32stubs.bas`.
- [x] UTF-8 string helpers in `mdGlobals.bas`: `ToUtf8Array`/`FromUtf8Array`
      plus `FromUtf8Ptr` (null-terminated UTF-8 `char*` → VB String, over
      `lstrlenA`+`MultiByteToWideChar`) — needed by `errmsg`/`column_text`.
- [x] `cConnection` core implemented (see section below).
- [x] ActiveX-DLL project (`TC6SQLiteTest`) wiring all 26 classes + modules
      together; class instancing restored (8 `MultiUse`, 18 `PublicNotCreatable`).
- [x] Class test runner driven over COM by `cTestHost` (PowerShell/VBScript),
      with test-name filter and output-file options; **15 tests / 69 checks
      PASS** against winsqlite3.dll (see the TestRunner section above).
- [x] `cRecordset` read core + `cFields`/`cField`; `cConnection.OpenRecordset`
      and `OpenSchema` wired to it (see section below).
- [x] Parameterised `cConnection.GetRs`/`ExecCmd` over `mdGlobals.BindVariant`
      (VB value → `bind_*`, `?` params, 1-based); `cRecordset` split into
      `pvPrepare`/`pvMaterialize` with a `frOpenParams` bind+load path.
- [x] `cCommand`/`cSelectCommand`/`cCursor` — reusable prepared statements
      with typed `Set*` binds and named-parameter lookup (see section below);
      statement helpers (`PrepareStatement`/`ReadColumnValue`/`StmtParam*`)
      shared in `mdGlobals`.
- [x] `cRecordset` batched write surface — field writes, `AddNew`/`Delete`/
      `UpdateBatch`/`ResetChanges`/`ContainsChanges`, PK-addressed write-back
      (see section below).
- [x] ~~Split `cRecordset` into a facade over an internal `cRowset`~~ —
      **rejected, not going to implement.** The proposal (strong `cRowset`
      references, orphaned fields stay readable) would diverge from the
      original: verified against RC6.dll 6.0.15 that an orphaned RC6 field
      raises exactly **err 91** on *every* recordset-dependent get/let
      (`Value`/`Name`/`ColumnType`/`ActualSize`/`Changed`/`OriginalValue`/
      `UnderlyingValue`/`Updateable`/…) while `IndexInFieldList` keeps
      working. The current weak-ref implementation (non-refcounted
      `m_pRs As cRecordset` — an early-bound call through the zeroed member
      raises 91 **before** any dereference) matches RC6 member-for-member
      and is kept as final; the two lifetime tests pin the contract.
- [x] Schema objects: `cDataBases`/`cDataBase` → `cTables`/`cTable` →
      `cColumns`/`cColumn` + `cIndexes`/`cIndex`, `cTriggers`/`cTrigger`,
      `cViews`/`cView`, wired from `cConnection.DataBases` (see section
      below).
- [x] `cRecordset` `Sort`/`SortRefresh` + `Find*` — in-memory multi-key
      sort and ADO-style criterion scan (see the cRecordset section).
- [x] `cField` metadata getters — origin db/table/column + decltype are now
      captured per column on **every** open (`pvCaptureColumnMeta`, also
      ReadOnly); the pragma-backed verdicts reuse the `cColumns` enrichment
      lazily, cached per origin table (`frColSchema`). `DefinedSize` parses
      the `(N)` of the declared type.
- [x] `cMemDB` — own `:memory:` connection on create (replaceable via
      `Set Cnn`), full delegation (`GetRs`/`Exec`/`ExecCmd`/commands/
      cursor/transactions), `GetTable` (WHERE/ORDER BY), `GetSingleVal`
      (Null on empty), `GetSum`/`GetAvg`/`GetMin`/`GetMax`/`GetCount`, and
      `CreateTableFromRs` for `cRecordset` sources (columns from field
      metadata incl. declared types and optional PKs; ADO/byte-content
      sources still raise). `cConnection.MemDB` returns a cached instance.
- [x] `cConnection.CreateTable`/`NewFieldDefs` — each FieldDefs item is a
      verbatim column definition (RC6's `NewFieldDefs` returns the
      intrinsic `Collection` — verified; `CreateTable` is not callable
      late-bound from VBScript, so the item syntax follows RC6's
      documented "name type" usage).
- [x] `cConnection.CopyDatabase` — `sqlite3_backup_init/step(-1)/finish`
      into a fresh connection, optional `VACUUM`; the copy is independent.
- [x] `UniqueID64` — format verified against RC6.dll 6.0.15: **local-time
      VB date serial × 10^14** (sub-ms bits from the high-res clock),
      returned as a **VT_I8 variant** (built by `mdGlobals.Int64Variant`
      via the Currency carrier; `BindVariant` accepts VT_I8 too).
      `mdGlobals.CreateUniqueID64` uses `GetSystemTimePreciseAsFileTime` +
      a strictly-increasing guard; `cRecordset.UniqueID64ToVBDate` is the
      exact inverse (+ sub-second remainder out-param). With
      `AutoCreateUniqueID64 = True` (default False, matches RC6) `AddNew`
      fills the INTEGER PK with a fresh id. DB NULLs (and `AddNew` cells)
      surface as **Empty** per `cConnection.MapDbNullToEmpty` — default
      True, `False` yields `Null` (probed against RC6 directly; the flag
      is copied onto the recordset at open time).
- [x] UDF/collation subsystem — `cConnection.Add/RemoveUserDefined*` over
      `create_function_v2`/`create_collation` with **plain VB6 AddressOf
      trampolines** in [src/mdUdf.bas](src/mdUdf.bas) (winsqlite3's whole
      ABI incl. callbacks is StdCall, so no thunks needed — verified by the
      tests). Each name from `DefinedNames` (comma-separated) gets a
      registry slot (object + ZeroBasedNameIndex); the 1-based slot number
      rides through SQLite as the user-data pointer. VB errors in callbacks
      never propagate into SQLite (converted to `sqlite3_result_error16`;
      collations fall back to "equal"). `cUDFMethods` wraps the live
      `sqlite3_context`: `Get*` argument readers (incl. VT_I8 `GetInt64`
      and ISO-date `GetDateTime`) and `SetResult*` setters (`TRANSIENT`
      buffers; `SetResultError` fails the statement — `pvMaterialize` now
      raises when the step loop ends in an error instead of returning a
      truncated result). Slots are released when their connection closes.
      `IFunction`/`IAggregateFunction`/`ICollation` stay empty by design —
      users Implements them.
- [x] `Content` Get/Let in the RC6-compatible blob format (byte-identical,
      cross-loads with the real RC6.dll both ways — see the format section
      below), `CreateTableFromRsContent` (default name from the blob's
      origin table, TEMP + WithPrimaryKeys options) and the
      `cMemDB.CreateTableFromRs` byte-content flavor over a shared
      `mdGlobals.CreateTableFromRecordset`.
- [x] `ContentChangesOnly` — changed-rows-only blob (see the format section
      below); loads via `Content` Let into a 0-row pending-ops recordset
      that `UpdateBatch` applies, cross-verified with RC6 in both
      directions.
- [x] `ToJSONUTF8` — byte-identical to RC6 (pinned by `JsonRC6Compat`, 12
      cases): `{"RecordCount": N,"Fields": [{ "Name", "Type" (decltype,
      empty for expressions), "PrimaryKey", "Nullable" (= Not NOT NULL),
      "DefaultValue" ("NULL" when none, empty for expressions)}],
      "RowsCols"/"ColsRows": [rows or transposed columns]}`. Values: null,
      numbers via `CStr(CDec(v))` (0.000015 not 1.5E-05; out-of-range
      doubles fall back to CStr → 1E+300), strings with `\" \\ \b \t \n
      \f \r` + lowercase `\uXXXX` for other control chars (and for >127
      with UniEscaping), blobs as Base64. Compact: one space after `:`,
      rows as `[ v,  v]`. Indent=N: line break (CRLF, or `Chr$(LfChar)`
      when nonzero) + N/2N-space levels, continuation lines align under
      the `"{ "`/`"[ "` opener (2N+2 spaces); empty recordset serializes
      `[]` inline in both modes.
- [x] `GetRowsWithHeaders` — GetRows plus a field-name header row/col at
      index 0 (index -1 with `HeaderAtIdxMinus1`, LBound -1); RowCount
      counts data rows only; TransPosed mirrors the shape.
- [x] ADO interop tail (all late-bound, no ADO reference; pinned against
      RC6 6.0.15 by mdConverterTests):
      - `cRecordset.GetADORsFromContent` — disconnected ADO recordset;
        INTEGER pk → adInteger, INTEGER → adDecimal, TEXT →
        adLongVarWChar, REAL → adDouble, BLOB → adLongVarBinary;
        attributes `adFldUpdatable|adFldIsNullable` + `adFldFixed` for
        numerics / `adFldLong` for text+blob / `adFldKeyColumn` for the
        pk; DB NULLs surface as ADO Null; int64 as Decimal.
        `cRecordset.DataSource` wraps the same recordset (RC6 returns a
        non-IDispatch binding object — divergence).
      - `cConnection.CreateTableFromADORs` — bracket-quoted CREATE TABLE
        from ADO field types (`AdoTypeToDeclType`: ints → INTEGER,
        boolean → BIT, floats/currency/decimal → REAL, dates → DATE,
        binary → BLOB, text → TEXT); single pk inline, composite as a
        table constraint; **no NOT NULL** on this path; RC6 quirk: a
        leading space after the paren when the table has no pk. Dates
        insert as `yyyy-mm-dd hh:nn:ss` text, booleans as 1/0.
      - `cConverter.ConvertDatabase/ConvertIndexes` — tables via
        `OpenSchema(adSchemaTables)`, pks via `adSchemaPrimaryKeys`,
        decltypes from `adSchemaColumns` (which adds `TEXT(n)` sizes from
        CHARACTER_MAXIMUM_LENGTH and ` NOT NULL` from IS_NULLABLE — the
        recordset-only path has neither); indexes via `adSchemaIndexes`,
        skipping pk indexes, named `[idx_<INDEX_NAME>]`. Progress events
        fire per table/row/index (RC6's exact event timing is not
        COM-observable — unverified). Storage-level compat pinned by
        quote()-dumps of both engines' converted files.
      - NB: RC6's recordset reader coerces values by decltype/width — now
        implemented in TC6, see "Reader value coercions" below.

- [x] Reader value coercions (probed against RC6 6.0.15, pinned by
      `ReaderCoercions` + the `date_bit` Content byte-compat case).
      Raw cells stay untouched; coercion happens **at access time**
      (`Fields().Value`, `ValueMatrix`, `GetRows*`) per column:
      - INTEGER columns type by their **load-time width class** (same
        rules as the Content W): W=1 → `Byte`, W=4 → `Long`, W=8 →
        VT_I8 **even for small values**; the pk column is always ≥ 4.
      - decltype containing DATE/TIME → `Date`; numeric storage is a
        **VB date serial** (not julian); unparsable text reads `Empty`.
      - decltype containing BIT/BOOL → `Boolean` = (value >= 1), so 2 →
        True but **-1 → False** and 0.5 → False.
      - `ToJSONUTF8` reads raw cells but serializes BIT/BOOL columns as
        JSON `true`/`false` literals and date columns as **quoted raw
        values** ("44562.4375"); Sort/Find/UpdateBatch/ChangesOnly all
        operate on the raw cells.
      - Content blob type codes: **6 = date** columns with cells stored
        as raw-value TEXT, **7 = boolean** columns with integer
        width-class cells; date columns count as text-ish for the
        ChangesOnly `textColIdx`.

## RC6 `Content` blob format (reverse-engineered, RC6.dll 6.0.15)

Decoded by differential probing (vary one datum, diff the blobs; probe
scripts in the session scratchpad). All numbers little-endian; strings are
`[Long cbBytes][UTF-16 chars]` (no terminator); per-row arrays allocate
**rows+1 elements** (one spare slot, presumably AddNew staging).

Implemented in `cRecordset.Content` Get/Let (`pvContentWrite`/
`pvContentRead`); **TC6-written blobs are byte-identical to RC6's** for the
probed scenarios (outside the two ignored pointer slots) and cross-load in
both directions — pinned by the `ContentRC6Compat` test which drives the
real RC6.dll via COM.

- **Header** (0x00): `Long cbTotal` (= blob size incl. itself), `Long
  RecordCount` ×3, `Long 0`, `Long CurRow` (−1 when empty), `Long
  4*(rows+1)`, `Long ptrGarbage`, `Long RecordCount`, `Long 0`, `Long
  ptrGarbage`, 5 zero `Long`s — the pointer slots are raw heap addresses
  dumped from RC6's internal structures and are ignored on load (verified:
  round-trip works across processes). Then the **row-order chain**: one
  `Long` per row holding the next row's index, natural order = `{1, 2,
  ..., N-1, 0}` (empty for 0 rows).
- **Strings**: DB filename (`:memory:`), the SQL text.
- **`Long FieldCount`**, then per field:
  - strings: db (`main`), `[table]`, `[column]` (bracketed), declared
    type, default value (`NULL` when none, else the text unquoted, e.g.
    `x` for `DEFAULT 'x'`), collation (`BINARY`), plain field name — all
    six metadata strings are **empty** for expression columns;
  - fixed block: `Long OriginTableIndex` (index into the tail's DML-set
    list, 0 for expressions), `Long ColumnOrdinal`, `Byte NotNull`,
    `Byte PK`, `Byte AutoInc` (UNIQUE is **not** serialized), `Long Type`
    (1=INTEGER, 2=FLOAT, 3=TEXT, 4=BLOB from the decltype; empty decltype
    → by the first value's storage class: text 3, blob 4, else 2 with
    values stored as doubles), `Byte 0`, `Long W` — for INTEGER columns
    the width class (1 = all values 0..255, 4 = int32, else 8; **the PK
    column is always ≥ 4**), else 0;
  - `[Long cb=rows+1][per-row not-null flag bytes + spare]`;
  - data: INTEGER `[Long cb=(rows+1)*W][values, W bytes each LE]` +
    `Long 0`; REAL `[Long cb=(rows+1)*8][doubles]` + `Long 0`; TEXT/BLOB
    `[Long cbContent][all content bytes back-to-back, text as UTF-8]`
    followed by an offsets array `[Long cb=(rows+1)*4][Long start offsets
    per row, final entry = cbContent]` (NULL rows have flag 0 and
    zero-length spans; no trailing `Long`).
- **Tail**: `Long cTables` = number of distinct origin tables
  (first-appearance order, 0 for expression-only selects), then per
  table: three DML template strings (`DELETE FROM [t] `, `UPDATE [t]
  SET `, `INSERT INTO [t] (`), `Long cKeyCols` + that many `Long`
  result-column ordinals — the table's PK columns as selected, or **all
  of its selected columns** when it has no PK. Then exactly **42 zero
  bytes**, and the whole blob is **padded to even length**.
- **Events**: RC6 raises no events on `Content` Let (verified via
  `WScript.ConnectObject`); TC6 matches (no `QueryFinished`).

NB: SQLite 3.42 (RC6) and newer engines parse some decimal literals
(e.g. `1e-300`) 1 ULP apart, so byte-compares must use binary-exact REAL
test values.

NB: engine versions (queried live via `SELECT sqlite_version()`): RC6.dll
6.0.15 embeds **3.42.0**; the VBSQLite objects in [cobj/](cobj/)
statically linked into the released TC6SQLite.dll are **3.51.1**. Those
two are the fixed, compatibility-relevant engines — engine diffs are what
surface test discrepancies. winsqlite3.dll (the dev/test path) is a
moving target that changes with Windows servicing, so its version is
deliberately not pinned anywhere.

NB: `cConnection.CreateTableFromRsContent` hit a **VB6 codegen bug** —
assigning a ByRef array *parameter* directly to a `Property Let` crashes at
runtime (0xC000008F); the parameter must be copied to a local array first.

NB: another **VB6 codegen hazard**, specific to the **default property**
(`VB_UserMemId = 0`): `cField.Value` calling a Friend method through a
Nothing typed reference raised error 91 but corrupted the heap once the
callee's prologue grew (frCellValue's coercion call) — crashing at process
teardown (0xC0000005, layout-sensitive). Identical non-default members
(`UnderlyingValue`/`OriginalValue` through the same zeroed weak ref, same
callee) raise a clean 91 — pinned by the orphan-field test. `cField.Value`
therefore guards `m_pRs Is Nothing` inline (hot path — kept inline).

### `ContentChangesOnly` blob format (changed rows only)

Same probing methodology (plus loader **mutation testing**: flip bytes in
an RC6-written blob, load + `UpdateBatch` in RC6, observe). The blob embeds
raw heap pointers and uninitialized union bytes, so byte-identity is not
attainable — the pinned contract is **semantic**: each engine loads the
other's blob via `Content` Let and `UpdateBatch` applies the pending ops
(covered by `ContentChangesOnlyRC6`). Layout:

- `Long -cbTotal` — **negative total size** marks the changes-only flavor
  (the `Content` Let reader dispatches on the sign), then `Long
  LoadedRows` (rows materialized at load — pending inserts and deletes do
  not change it), DB filename, SQL.
- `Long FieldCount` + per field the same seven metadata strings and fixed
  block as the full blob, but **no per-row arrays / cell data**; the
  INTEGER width class W keeps its **load-time value** (RC6 never widens it
  for pending edits — original values are used, pending inserts excluded).
- The same DML-template tail (sets + key ordinals), **without** the
  trailing zero block.
- **Changed cells**: `Long cCells`, then per cell a `Long` slot id —
  numeric-valued cell `(label << 9) Or colOrdinal`; pointer-valued cell
  (text/blob) `Not ((label << 9) Or textColIdx)` where `textColIdx`
  counts the table's **non-key TEXT/BLOB columns**; deleted row
  `label << 9`. `label` is the row's **stable physical id** (0-based load
  position; survives deletes/sorts, `AddNew` allocates the next one —
  mirrored by `m_aRowLabel`/`m_lNextLabel`), and cells emit sorted by
  label. Then 16 zero bytes (ignored), then per cell
  `[Long vt][Long junk][8-byte union]`: vt 0 = NULL (union garbage), vt 2
  = Integer in the low 2 bytes (rest garbage), vt 3 = Long inline, vt 5 =
  Double, vt 14 = int64 as the Decimal mantissa low 8 bytes; pointer
  cells vt 3 with `[Long cbBytes][ptr]` — **negative cb = UTF-16 text,
  positive cb = raw blob bytes** — content concatenated in a **shared
  stream right after the records**. A deleted row contributes one vt 6
  cell with the **magic Currency 987654321098.765** (raw int64
  &H2316A9E9B32082); an insert carries only its **assigned** cells except
  an unassigned pk, which gets a second magic (raw &H2316A9E9B318C6 =
  "let SQLite assign"). Junk fields and pointers are ignored by the
  loader (mutation-verified).
- **Rows needing a WHERE** (updates + deletes, not inserts): `Byte 0`,
  `Long cRows`, then a **flag byte — 1 when a text-pk WHERE row carries
  label 0, else 0** — then per row a `Long` handle — numeric pk
  `2 * label`; text pk `&H800000` for label 0 else `&H1000000 - 2 *
  label` (RC6's internal sorted-key handles; its loader **validates**
  these, so the formulas must be reproduced exactly). Then 15 zero bytes,
  then per row `[Long vt][Long junk][8-byte union]` with the pk WHERE
  value (Long inline, or `[-cb][ptr]` with the text in a second stream
  after the records), a final zero byte and even-length padding.
- The WHERE targeting comes from the pk value; the labels join cells to
  rows and (for int pk) only need mutual consistency (mutation-verified),
  but reproducing RC6's physical labels keeps the blobs structurally
  identical — pinned by `ChangesOnlyRC6Compat`, which drives every
  scenario through both engines and compares the blobs byte-for-byte
  outside the junk/pointer fields, then cross-applies each blob in the
  opposite engine.
- Loading a changes blob yields a 0-row recordset with `ContainsChanges` =
  True and `Updatable` = False (RC6-verified); `UpdateBatch` then executes
  DELETE/UPDATE/INSERT against the attached connection addressed by the
  single key column from the DML tail (composite-pk change blobs are not
  supported — RC6's writer emits one pk record per row regardless).

- [ ] Tail / possibly out of scope for v1: DDL-parsing members
      (`cColumn.OriginalConstraint`/`ConstraintName`/`CheckExpression`/
      `PrimarySortOrder`, `cTable.Constraint`), `cCommand`/
      `cSelectCommand.Save` + `Repl*`, ADO interop (`cConverter`,
      `CreateTableFromADORs`, `DataSource`). `cDBAccess` stays a stub
      permanently (user decision).

## [src/cConnection.cls](src/cConnection.cls) — connection wrapper

First class fleshed out. A thin wrapper over the `stub_sqlite3_*` declares holding the
`sqlite3*` handle (`m_hDb As LongPtr`) and a `VBA.Collection` savepoint stack.

**Implemented and verified** (call sequences exercised at module level against
winsqlite3.dll — open/exec/pragma-scalar/savepoints/errmsg/rowid all
pass):

- Lifecycle: `CreateNewDB`/`OpenDB` (`READWRITE|CREATE`), `OpenDBReadOnly`
  (`READONLY`) via `sqlite3_open_v2`; `Class_Terminate` closes with
  `close_v2`. `EncrKey`/`EnableVBFunctions` args accepted but ignored (no
  codec / no VB-UDFs yet); `ReKey` raises (winsqlite3 has no codec).
- `Execute` (via `sqlite3_exec`, raises `vbObjectError` with `errmsg` on
  failure); `AttachDataBase`/`DetachDataBase`/`CompactDataBase` as DDL.
- Transactions: `BeginTrans`/`CommitTrans`/`RollbackTrans` maintain the
  savepoint stack (empty-string marker = outer `BEGIN`; a name = `SAVEPOINT`),
  emit `BEGIN`/`COMMIT`/`ROLLBACK` and `SAVEPOINT`/`RELEASE`/`ROLLBACK TO`,
  and raise the matching events. Nesting gated on `EnableNestedTransactions`.
- Info/pragmas: `Version`, `LastDBError`/`GetSqliteErrStr`/`LastDBErrCode`,
  `AffectedRows`/`TotalAffectedRowsInSession`, `LastInsertAutoID`,
  `PageSize`/`Synchronous`/`BusyTimeOutSeconds`, `CheckIntegrity`, `Cancel`
  (`sqlite3_interrupt`). Identifiers/strings quoted via private `pvQuoteId`/
  `pvQuoteStr` (private helpers carry the `pv` prefix per CLAUDE.md).
- Date helpers (`GetDateString` etc.) as SQLite text formats; `NewFieldDefs`
  returns `New Collection`.
- **int64 on x86**: `LastInsertAutoID` recovers the true rowid from the
  `Currency` the API returns (raw int64 bits) via `CDec(cur) * 10000`
  (`#If Win64` uses the native `LongLong`). Verified: rowids 1, 2.

`OpenRecordset`/`OpenSchema` build a `cRecordset` via the `Friend Property Get
frDbHandle` accessor that hands the raw `sqlite3*` to the recordset.
`GetRs(SQL, ...params)` and `ExecCmd(SQL, ...params)` bind `?` parameters
through `mdGlobals.BindVariant` (VB value → `bind_int`/`bind_double`/
`bind_text` UTF-8/`bind_blob`/`bind_null`, `SQLITE_TRANSIENT` so SQLite copies
the buffer); `GetRs` runs `cRecordset.frOpenParams` (prepare→bind→materialise),
`ExecCmd` prepares/binds/steps to `SQLITE_DONE`, both finalising on any error.
Two mappings that matter: **integral `Currency`/`Decimal` go through
`bind_int64`** (a double round-trip would silently lose precision above 2^53;
on x86 the declare types the int64 as `Currency`, so the value is scaled down
by 10000 to land in the raw bits), and **`Date` binds as ISO text**
(`yyyy-mm-dd hh:nn:ss`, matching `GetDateString` — as a date-serial double it
would never equality-match text-stored dates).

`CreateCommand`/`CreateSelectCommand`/`CreateCursor` build the corresponding
prepared-statement wrappers (below).

`DataBases` returns a fresh `cDataBases` snapshot (see the schema-objects
section).

**Left as stubs** (need classes not yet built): `MemDB`, `CopyDatabase`, the
UDF/collation add/remove pair, and the ADO/`CreateTableFrom*` migration
surface.

## Schema objects — cDataBases → cTables → cColumns / cIndexes / cTriggers / cViews

Read-only schema tree over pragmas + `sqlite_master`, built **on top of the
recordset machinery** (`cConnection.GetRs`). Every collection is a snapshot
taken at creation, and every parent property (`cDataBase.Tables`,
`cTable.Columns`, …) builds a fresh snapshot on access — nothing is cached,
so `ReScanSchemaInfo` is a no-op. All collections follow the `cFields`
pattern: `VBA.Collection` keyed on name (O(1) `Item(name)`), `Item` accepts a
0-based index or a name, `NewEnum` enables `For Each`.

- `cConnection.DataBases` → `cDataBases` (`PRAGMA database_list`), with
  `AttachDataBase`/`DetachDataBase` delegating to the connection and
  re-snapshotting. `cDataBase`: `Name`/`NameInBrackets`/`Tables`/`Views`.
- `cTables` (per database): `sqlite_master WHERE type='table'`, internal
  `sqlite_*` tables excluded. `cTable`: `Name`, `SQLForCreate` (the stored
  DDL), `Columns`, `Indexes`, `Triggers`; `Constraint` (table-level clause)
  needs DDL parsing — not implemented.
- `cColumns` (per table): `PRAGMA table_info` enriched with
  `sqlite3_table_column_metadata` (collation sequence + autoincrement) and a
  single-column-UNIQUE-index scan (`PRAGMA index_list`/`index_info`,
  `unique<>0`). `cColumn`: `Name`, `ColumnType` (declared type),
  `DefaultValue` (verbatim expression text, e.g. `'x'`),
  `NotNullConstraint`, `PrimaryKey`/`PrimaryAutoIncrement`, `Collate`,
  `UniqueConstraint`. The DDL-text-only members (`OriginalConstraint`/
  `ConstraintName`/`CheckExpression`/`PrimarySortOrder`/conflict
  algorithms) return defaults — they need CREATE TABLE parsing.
- `cIndexes` (per table): `sqlite_master WHERE type='index'`, including the
  implicit `sqlite_autoindex_*` entries (whose `SQL` is empty).
  `cTriggers` (per table) and `cViews` (per database) likewise; `cView.SQL`
  is the best-effort SELECT body (text after the first ` AS ` in the DDL).

## [src/cCommand.cls](src/cCommand.cls) / cSelectCommand / cCursor — prepared statements

Three thin wrappers over one prepared `sqlite3_stmt*` each, created by the
`cConnection` factories; all share the `mdGlobals` statement helpers
(`PrepareStatement`, `BindInt64Value`/`BindTextValue`/`BindBlobValue`,
`ReadColumnValue`, `StmtParamIndex`/`StmtParamName`). Each holds a **strong**
connection reference (child → parent only, no cycle) and finalizes its
statement in `Class_Terminate`; the `SQL` Let re-prepares.

- `Save CommandKey` (probed against RC6 6.0.15): persists the ORIGINAL
  template SQL as a **UTF-16 blob** into `dhWriteCommands` (cCommand) /
  `dhSelectCommands` (cSelectCommand), created on demand with RC6's exact
  DDL — `(ID Integer Primary Key,CommandKey Text Collate NoCase Unique On
  Conflict Replace, SQL Blob)`. `CreateCommand`/`CreateSelectCommand`
  resolve a saved key (case-insensitive) to its stored SQL before
  treating the argument as SQL.
- `ReplColumnOrTableName`/`ReplTextBlock` (cSelectCommand): replace the
  Nth bare `?` (string/identifier literals and comments skipped) in the
  **executed** SQL — identifiers bracket-quoted, text blocks raw — and
  re-prepare; the `SQL` property keeps returning the original template
  (RC6-verified), bindings are lost, remaining `?`s renumber. Divergence:
  RC6 6.0.15 breaks when a non-last `?` is replaced and params are bound
  afterwards ("incomplete input" / error 438) — TC6 handles it correctly.
- Statements can't prepare against a `?` in table position, so table
  replacement only works where the template stays parseable — same
  limitation as RC6 (its `SELECT ... FROM ?` create also fails).

- Common `Set*` surface (1-based param index): `SetText` (UTF-8),
  `SetTextPtr` (UTF-16 ptr → `bind_text16`), `SetTextUTF8Ptr`, `SetBlob`/
  `SetBlobPtr`, `SetInt32`, `SetInt64` (Variant carrier → `bind_int64`),
  `SetDouble`, `SetDate`/`SetShortDate`/`SetTime` (ISO text), `SetBoolean`
  (0/1), `SetNull`, `SetAllParamsNull` (`clear_bindings`). All ptr/text/blob
  binds use `SQLITE_TRANSIENT`.
- Named params: `NameToIdx` (tries the bare name, then `:`/`@`/`$` prefixes),
  `IdxToName` (returns the prefixed name), `ParameterCount`.
- `cCommand.Execute` steps to `SQLITE_DONE` then **resets**, so the command
  can be re-bound and re-run; raises with `errmsg` on failure.
- `cSelectCommand.Execute` materialises a disconnected `cRecordset` from the
  live statement via `cRecordset.frLoadStmt` (reads to completion **without**
  finalizing — the command owns the statement) and then resets it; the
  returned recordsets outlive later Executes independently.
- `cCursor`: forward-only `Step` (True on row, False on done, raises on
  error, increments `StepCounter`), `Reset` (rewind + zero counter), and
  per-column `ColCount`/`ColName`/`ColType`/`ColVal` reading the current row
  directly off the statement.
- `Save(CommandKey)` and `cSelectCommand.ReplColumnOrTableName`/
  `ReplTextBlock` raise `Not implemented yet`.

## [src/cRecordset.cls](src/cRecordset.cls) — disconnected recordset

Reads the **entire** result set into an in-memory matrix on open, then serves
navigation/field access from memory (RC6's recordsets are disconnected). Built
by `cConnection.OpenRecordset`/`OpenSchema` (`frOpen`), or via the public
`OpenRecordset(SQL, Cnn, ReadOnly)`.

- **Load**: `prepare_v2` → `step` loop → `column_*` into `m_vData(col, row)`
  (a `Variant` matrix, last dim = row so `ReDim Preserve` can grow it by
  doubling). Column values map by storage class: INTEGER via the x86
  `Currency`→`CDec*10000` int64 recovery (Long when it fits, else Decimal),
  FLOAT→`Double`, TEXT→`FromUtf8Ptr`, BLOB→`Byte()` (`RtlMoveMemory` from
  `column_blob`), NULL→`Null`.
- **Navigation**: `RecordCount`, `BOF`/`EOF`, `MoveFirst/Last/Next/Previous`,
  `AbsolutePosition`/`Bookmark` (get/set), `ReQuery`; raises `Move`/
  `QueryFinished`. Empty set ⇒ `BOF And EOF`.
- **Fields**: `Fields` (default member) → `cFields`, a collection of `cField`
  bound to the recordset's *current* row (`frInit`/`frCellValue`). `cField`:
  `Value` (Get), `Name`, `ColumnType` (inferred from the VB value),
  `IndexInFieldList`, `ActualSize`. `cFields`: `Item` (0-based index **or**
  name), `Count`, `Exists`, `NewEnum` (`For Each`). The backing
  `VBA.Collection` is **keyed on field name**, so `Item(name)`/`Exists` are
  O(1) (no per-access scan) — duplicate result-column names keep the first
  occurrence keyed, the rest reachable by index.
- **Read helpers**: `ValueMatrix(row, col)` (Get), `GetRows` (ADO-style
  `(field, record)` array; `TransPosed` and a comma-separated field subset
  supported).
- **No reference cycle**: the recordset owns `cFields`, which owns the
  `cField`s. Their back-references to the recordset are **weak**: `m_pRs` is
  typed `As cRecordset` but the raw pointer is `CopyMemory`'d in/out (never
  `Set`), so no AddRef/Release ever fires and members call through it
  directly with no per-access accessor cost. `Class_Terminate` re-zeroes it
  the same way — otherwise VB's teardown would auto-Release the still-alive
  recordset through the non-refcounted member (unbalanced Release). A strong
  back-ref would be a cycle that keeps the recordset (and
  its whole result matrix) alive forever.
  This weak-ref design is **final**: it matches the original RC6, where an
  orphaned field raises err 91 on every recordset-dependent member while
  `IndexInFieldList` keeps working (verified against RC6.dll 6.0.15). A
  facade + `cRowset` split with strong references was considered and
  rejected — it would diverge from RC6 by keeping orphaned fields readable.
- **Batched write surface**: opening a recordset analyses updatability —
  every table-backed result column must originate in **one** table
  (`column_table_name`/`column_origin_name`/`column_database_name` statement
  metadata) and all of that table's PK columns (`PRAGMA table_info`, `pk`
  ordinal) must be present in the result; `Updatable`/`cField.Updateable`
  expose the verdict (expression columns are never writable). Field writes
  (`cField.Value` Let, `ValueMatrix` Let) capture the cell's original on
  first change; `AddNew` appends an all-`Null` row; `Delete` removes the
  current row keeping a full snapshot. `UpdateBatch` writes everything back
  inside a `tc6_updatebatch` savepoint — `DELETE`/`UPDATE` (dirty columns
  only) address rows by `WHERE pk = ?` using **original** PK values,
  `INSERT` sends all table-backed columns (a `Null` single-column INTEGER
  PK auto-assigns; the new rowid is backfilled into the cell via
  `last_insert_rowid`) — and rolls back to the savepoint on any error.
  `ResetChanges(ResetAll/DeletesOnly/InsertsOnly/UpdatesOnly)` discards
  pending changes (restoring deleted-row snapshots); `ContainsChanges` and
  `cField.Changed`/`OriginalValue` report pending state.
- **Sort / Find**: `Sort = "col1 [DESC], col2 …"` (also `[bracketed]` names)
  applies immediately and `SortRefresh` re-applies after data changes — a
  stable multi-key quicksort over a row-index array, then one permute pass
  that moves the matrix **and** the pending-changes state together; the
  cursor follows its row (raising `Move`). Value ordering follows SQLite:
  NULL < numeric (`CDec` compare) < text (binary) < blob. `FindFirst`/
  `FindNext`/`FindPrevious`/`FindLast` scan with an ADO-style criterion:
  `field op literal` with ops `=`, `<>`, `<`, `<=`, `>`, `>=`, `LIKE`
  (SQLite `%`/`_` wildcards, case-insensitive) and `IS [NOT] NULL`;
  literals are `'strings'` (`''` escape), numbers (locale-independent
  `Val`) or `NULL`. On a hit the cursor moves there (`Move` event) and True
  returns; on a miss the position stays. `DistinctNullValues:=True`
  (default) = SQL three-valued logic (`= NULL` never matches); `False`
  treats NULLs as regular equal values (`= NULL` matches null cells, `<>`
  counts them).
- **Left raising `Not implemented`**: `Content*`, `ToJSONUTF8`,
  `GetRowsWithHeaders`, ADO interop.

**Gotcha:** `cField`'s `FieldType` enum (`SQLite_INTEGER`…) collides
case-insensitively with `sqlite3win32helper`'s `SQLITE_INTEGER`… constants. VB6 only
errors on *unqualified* use from a third module — `cRecordset.pvReadColumn`
qualifies them `sqlite3win32helper.SQLITE_*`; test code avoids both names.
