# RC6 SQLite Replacement — Working Notes

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

## [test/](test/) — VB6 compile & run harness

A Std-EXE VB6 project that compiles `src\mdSqliteApi.bas` + `src\mdGlobals.bas`
and exercises them at runtime, built and run headless with the OS's `VB6.EXE`
(`/make ... /out`).

- [test/apitest.vbp](test/apitest.vbp) — Std-EXE, `Startup="Sub Main"`. Uses the
  full field set VB6 writes (Command32, MajorVer, CompilationType, …); a stripped
  vbp made `/make` mis-parse the module list. **Files must be CRLF** — LF made
  VB6 drop modules ("Must have startup form or Sub Main()").
- [test/mdMain.bas](test/mdMain.bas) — `Sub Main` opens `:memory:`, prepares
  `SELECT 40+2`, steps, reads the column, closes; writes results to
  `apitest_out.txt`. Includes a `LongPtr`→ANSI-string helper.

Result (proves the StdCall declares are binary-correct — a wrong convention
would corrupt the stack, not compute 42):

```
libversion=3.51.1
open rc=0 hDbNonZero=True
prepare rc=0
step rc=100 (100=SQLITE_ROW) col0=42
close rc=0
```

Build/run recipe (headless): `Start-Process VB6.EXE -ArgumentList '/make',
<vbp>,'/out',<log> -Wait` — VB6 is a GUI-subsystem app, so the shell must wait
on it explicitly, and the `/out` log is **append-mode**.

## [test/TestRunner.vbp](test/TestRunner.vbp) — class test runner

A single Std-EXE that compiles **all 26 classes + both shared modules** and runs
a suite of tests against them, so the replacement can be developed test-first.
Built/run headless like `apitest`. Current run: **8 tests / 33 checks, all
PASS** (exercises the real `cConnection` against winsqlite3.dll 3.51.1).

- [test/mdTestRunner.bas](test/mdTestRunner.bas) — the harness. `Sub Main` calls
  one `RunXxxTests` per class; `TestBegin`/`TestEnd`/`TestErr` bracket each test
  (own `On Error GoTo EH`, so one failure never aborts the rest);
  `AssertTrue`/`AssertEqLng`/`AssertEqStr` record pass/fail and keep going.
  Output goes to `test\testrun.log`.
- [test/mdConnectionTests.bas](test/mdConnectionTests.bas) — one standard module
  **per class** holds that class's tests (`Test_*` subs + a public
  `RunConnectionTests`). Add a new module + one line in `Main` for each class.

Two things the runner project needs that `apitest` did not:

- **`Reference=…stdole2.tlb#OLE Automation`** in the vbp — the classes use
  `IUnknown` (the `NewEnum` collection members); without the OLE Automation
  reference VB6 reports *"User-defined type not defined"*.
- **Class instancing = Private.** A Standard EXE only permits `Private` classes
  (VB6 force-downgrades anything public with a warning). All 26 `.cls` were set
  to `MultiUse = 0 / VB_Creatable = False / VB_Exposed = False`. This also fixed
  a latent bug: the schema/child stubs had an **invalid** header (`MultiUse = -1`
  with `VB_Creatable = False`) that loaded in no project type. **When the
  ActiveX-DLL project is built, restore instancing**: the 8 creatable classes
  (`cConnection`, `cRecordset`, `cMemDB`, `cConverter`, `cDBAccess`, `IFunction`,
  `IAggregateFunction`, `ICollation`) → `MultiUse`; the other 18 →
  `PublicNotCreatable`. `New` still works on `Private` classes within the project,
  so the tests are unaffected.

## Status / next steps

- [x] SQLite classes identified from the typelib.
- [x] All 26 classes stubbed with full signatures, empty bodies.
- [x] SQLite backend binding: `mdSqliteApi.bas` with all winsqlite3.dll exports
      declared (StdCall, verified) plus core constants; UTF-8 helpers in
      `mdGlobals.bas`.
- [x] VB6 (x86) + tB dual-target compatibility of `mdSqliteApi.bas`.
- [x] Std-EXE test harness compiles AND runs against winsqlite3.dll
      (open/prepare/step/column/close all verified, SQLite 3.51.1).
- [x] UTF-8 string helpers in `mdGlobals.bas`: `ToUtf8Array`/`FromUtf8Array`
      plus `FromUtf8Ptr` (null-terminated UTF-8 `char*` → VB String, over
      `lstrlenA`+`MultiByteToWideChar`) — needed by `errmsg`/`column_text`.
- [x] `cConnection` core implemented (see section below).
- [x] Class test runner: single Std-EXE (`TestRunner.vbp`) with all 26 classes
      + shared modules + one test module per class; 8 tests / 33 checks PASS
      against winsqlite3.dll (see the TestRunner section above).
- [ ] Implement `cRecordset` (query/navigation), then the object-factory
      members of `cConnection` that return it (`OpenRecordset`/`GetRs`/
      `OpenSchema`) and `cCommand`/`cSelectCommand`/`cCursor`.
- [ ] Add a `.twinproj`/project file wiring the classes together. All class
      references are now project types or intrinsics (`cCollection` → the
      intrinsic `VBA.Collection`), so an ActiveX-DLL project no longer needs
      an external RC6 reference to compile.

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

**Left as stubs** (need classes not yet built): `OpenRecordset`/`GetRs`/
`OpenSchema`/`ExecCmd`, `CreateCommand`/`CreateSelectCommand`/`CreateCursor`,
`DataBases`/`MemDB`, `CopyDatabase`, the UDF/collation add/remove pair, and
the ADO/`CreateTableFrom*` migration surface.
