# TC6 — SQLite for VB6 and TwinBasic

TC6 reimplements the SQLite subsystem of Olaf Schmidt's [vbRichClient6](https://vbrichclient.com)
(RC6) as a plain VB6 ActiveX DLL, so no third-party binaries need to be
shipped or registered: in source form the API declares bind to the
OS-supplied `winsqlite3.dll` (present on every Windows 10/11 box), while
the release DLL embeds the SQLite routines outright (see Building below)
and runs on Windows XP too. The public interface follows RC6's
`cConnection`/`cRecordset`/`cCommand` object model closely enough that
most RC6 database code runs unmodified, including binary interop:
`cRecordset.Content` blobs are byte-identical to RC6's and load in either
library.

## Layout

| Path | Contents |
|------|----------|
| [src/](src/) | Library sources + [TC6SQLite.vbp](src/TC6SQLite.vbp) (ActiveX DLL, binary compatibility vs `TC6SQLite.cmp`) |
| [test/](test/) | Test suite + [TestRunner.vbp](test/TestRunner.vbp) (ActiveX DLL driven over COM) |
| [cobj/](cobj/) | Prebuilt object files swapped in at link time by `replace_cobj.bat` |
| [doc/](doc/) | RC6 typelib IDL and SQLite amalgamation for reference |
| [NOTES.md](NOTES.md) | Living implementation notes, incl. the reverse-engineered `Content` blob format |

### Classes

`cConnection`, `cRecordset` (read + batch updates, RC6-compatible `Content`
serialization), `cFields`/`cField`, `cCommand`/`cSelectCommand` (prepared
statements), schema objects (`cDataBases`/`cTables`/`cColumns`/`cIndexes`/
`cTriggers`/`cViews`), `cMemDB`, and the UDF/collation subsystem
(`IFunction`/`IAggregateFunction`/`ICollation` implemented by user classes,
registered via `cConnection.AddUserDefined*`).

## Building

Requires VB6 (SP6). [TC6SQLite.vbp](src/TC6SQLite.vbp) must be compiled
**from the VB6 IDE** with the custom linker add-in loaded — a command-line
`VB6.EXE /make` build appears to succeed but does not produce a correct
DLL. The test project has no such requirement and builds headlessly with
`LIB` pointing at the VC98 link libraries:

```bat
set LIB=C:\Program Files (x86)\Microsoft Visual Studio\VC98\Lib;C:\Program Files (x86)\Microsoft Visual Studio\VB98
VB6.EXE /make test\TestRunner.vbp /out build_test.log
```

When built with the add-in loaded, `replace_cobj.bat` runs before linking
and replaces the compiled objects of the `sqlite3win32stubs.bas`/
`sqlite3win32helper.bas` modules with the prebuilt `.obj` files from
[cobj/](cobj/) — thanks to the [VBSQLite](https://github.com/Kr00l/VBSQLite)
project. These statically link the SQLite routines into the final DLL, so
the release binary does not use `winsqlite3.dll` at all and works on
Windows XP as well. A headless `/make` (as used for the test DLL above)
skips the swap: the resulting binary calls `winsqlite3.dll` through the
declares, exactly as the code does when run directly in the VB6 IDE.

## Running the tests

The suite is an ActiveX DLL (`TC6SQLiteTest.cTestHost`) driven by a 32-bit
script host; the process exit code is the failed-test count:

```bat
C:\Windows\SysWOW64\cscript.exe //nologo test\run_tests.vbs [/out:log-file] [/filter:tokens]
```

The `ContentRC6Compat` test cross-checks blob compatibility against the real
RC6.dll and expects it registered on the machine; it is skipped otherwise.

## Compatibility notes

- Development and tests bind to the winsqlite3.dll shipped with Windows;
  the release DLL links the embedded VBSQLite objects instead. No
  encryption codec, so `ReKey`/`EncrKey` raise/no-op.
- DB NULLs surface as `Empty` by default, `cConnection.MapDbNullToEmpty =
  False` restores `Null` (matches RC6).
- 64-bit integers use `Variant` of `VT_I8`, `UniqueID64` values match RC6's
  local-time encoding.
- `Content`/`ContentChangesOnly` blobs, `ToJSONUTF8` output, event timing
  and the ADO interop (`GetADORsFromContent`, `CreateTableFromADORs`,
  `cConverter`) are pinned against the real RC6.dll by the test suite.
- `cDBAccess` (RC6 thread marshalling) is intentionally left unimplemented;
  see the checklists in [NOTES.md](NOTES.md).
