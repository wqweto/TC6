# TC6 SQLite Replacement — Working Notes

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

## [src/mdSqliteApi.bas](src/mdSqliteApi.bas) — winsqlite3.dll API layer

(originally `mdWinApi.bas`; renamed once it held only the `vbsqlite3_*` declares —
the UTF-8 helpers moved to [src/mdGlobals.bas](src/mdGlobals.bas).)

Shared standard module holding the raw SQLite API surface. Generated by dumping
the export table of the OS-supplied `C:\Windows\System32\winsqlite3.dll`
(SQLite 3.x, **297** exports) and parsing the prototypes from
[doc/sqlite3.h](doc/sqlite3.h).

- **283** functions declared (`Public Declare … Lib "winsqlite3"`), later reduced
  to 275 (StdCall-only). Each is named `vbsqlite3_Xxx` with `Alias "sqlite3_Xxx"`
  — the `vb` prefix avoids clashing with any real `sqlite3_*` symbol and keeps
  call sites greppable, while the Alias binds to the true DLL export.
- **14** exports are data symbols or internal helpers with no public prototype
  (`sqlite3_version[]`, `sqlite3_temp_directory`, `sqlite3_data_directory`,
  `sqlite3_fts3_may_be_corrupt`, `sqlite3_unsupported_selecttrace`, and the
  internal `sqlite3_win32_*` helpers) — listed in a footer comment, declare when
  needed.
- Only a **curated core set of constants** is included (result codes, datatypes,
  text encodings, open/prepare flags, `SQLITE_STATIC`/`SQLITE_TRANSIENT`). The
  full constant set stays in `doc\sqlite3.h` per the instruction.

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
  `WideCharToMultiByte`/`MultiByteToWideChar` with `CP_UTF8`. `mdSqliteApi.bas`
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
**8 tests / 33 checks, all PASS** against winsqlite3.dll 3.51.1.

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
- [x] SQLite backend binding: `mdSqliteApi.bas` with all winsqlite3.dll exports
      declared (StdCall, verified) plus core constants; UTF-8 helpers in
      `mdGlobals.bas`.
- [x] VB6 (x86) + tB dual-target compatibility of `mdSqliteApi.bas`.
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
- [ ] `cRecordset` write surface (AddNew/Delete/UpdateBatch/ResetChanges),
      Sort/Find, Content serialization, ADO interop, JSON export.
- [ ] Parameterised `cConnection.GetRs`/`ExecCmd` and
      `cCommand`/`cSelectCommand`/`cCursor` (need value→`bind_*` binding).

## [src/cConnection.cls](src/cConnection.cls) — connection wrapper

First class fleshed out. A thin wrapper over `mdSqliteApi` holding the
`sqlite3*` handle (`m_hDb As LongPtr`) and a `VBA.Collection` savepoint stack.

**Implemented and verified** (call sequences exercised at module level against
winsqlite3.dll 3.51.1 — open/exec/pragma-scalar/savepoints/errmsg/rowid all
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

`OpenRecordset`/`OpenSchema` are now implemented (they build a `cRecordset`,
via the `Friend Property Get frDbHandle` accessor that hands the raw
`sqlite3*` to the recordset).

**Left as stubs** (need classes/binding not yet built): `GetRs`/`ExecCmd`
(parameterised), `CreateCommand`/`CreateSelectCommand`/`CreateCursor`,
`DataBases`/`MemDB`, `CopyDatabase`, the UDF/collation add/remove pair, and
the ADO/`CreateTableFrom*` migration surface.

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
  `cField`s. Their back-references to the recordset are **weak** (raw `ObjPtr`,
  dereferenced via `mdGlobals.ObjectFromPtr` — `__vbaObjSetAddref` from a
  pointer). A strong back-ref would be a cycle that keeps the recordset (and
  its whole result matrix) alive forever. Proven by the `NoReferenceCycle`
  test: a module-level `g_lLiveRecordsets` (bumped in `Class_Initialize`/
  `Terminate`) returns to baseline after a recordset goes out of scope.
- **Left raising `Not implemented`**: the write surface (`AddNew`/`Delete`/
  `UpdateBatch`/`ResetChanges`/`ValueMatrix` set/`cField.Value` set), `Sort`/
  `Find*`, `Content*`, `ToJSONUTF8`, `GetRowsWithHeaders`, ADO interop.

**Gotcha:** `cField`'s `FieldType` enum (`SQLite_INTEGER`…) collides
case-insensitively with `mdSqliteApi`'s `SQLITE_INTEGER`… constants. VB6 only
errors on *unqualified* use from a third module — `cRecordset.pvReadColumn`
qualifies them `mdSqliteApi.SQLITE_*`; test code avoids both names.
