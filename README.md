# modScriptApdObjects

VBA module for Access 2003 ADP projects (Access Data Project). Exports all
client-side "scriptable" objects (Forms, Reports, Macros, Modules) as text
files and can import them back into an ADP. Intended for version control
(Git/SVN) and diffing of legacy ADP projects.

**Version:** see the `MODULE_VERSION` constant in the module (format `Major.Minor.yyyymmdd`)

## What gets exported

| Object type | Folder       | File extension |
|-------------|--------------|-----------------|
| Forms       | `Forms\`     | `.frm`          |
| Reports     | `Reports\`   | `.rpt`          |
| Macros      | `Macros\`    | `.mac`          |
| Modules     | `Modules\`   | `.bas`          |

Export is done via `Application.SaveAsText`.

**Not included:** Server-side objects (Views, Stored Procedures, Functions)
live in the SQL Server backend, not in the ADP file, and are not handled by
this module.

## Installation

1. Import the module `modScriptApdObjects.bas` into the ADP
   (VBA editor → File → Import File…).
2. Done — no additional references required.

## Export

### Interactive

```vba
RunExportAllScriptableObjects
```

Prompts for a target path via an input box. Leaving the field empty uses
the default path (see below).

### From code

```vba
' Default path (adp_dump next to the ADP file), no log file
ExportAllScriptableObjects

' Custom target path, no log file
ExportAllScriptableObjects "C:\Export\MyProject\"

' Custom target path, with log file
ExportAllScriptableObjects "C:\Export\MyProject\", WithLogFile:=True

' Default path, with log file
ExportAllScriptableObjects WithLogFile:=True
```

**Parameters:**

| Parameter      | Type    | Optional | Default | Description |
|----------------|---------|----------|---------|-------------|
| `ExportFolder` | String  | Yes      | `""`    | Target directory. If empty, `<ADP's directory>\adp_dump\` is used. |
| `WithLogFile`  | Boolean | Yes      | `False` | `True` additionally creates a log file `ExportLog_yyyymmdd_hhnnss.txt` in the target directory with details for every exported object plus success/failure counts. |

### Cleanup behavior before export

Before each run, **only the four object subfolders** (`Forms`, `Reports`,
`Macros`, `Modules`) are recursively cleared, so stale files from
deleted/renamed objects don't linger. The rest of the target directory is
left untouched, e.g.:

- Log files in the base directory
- A `.git` folder (Git repository)
- A `.svn` folder (SVN working copy)

`.git` and `.svn` folders are also never deleted even if they happen to
exist **inside** one of the four object subfolders.

At the end of the export a summary is shown (number succeeded/failed, and
the log file path if applicable).

## Import

Loads a single previously exported file back into the currently open ADP:

```vba
' Object name is derived from the file name -> "frmCustomer"
ImportScriptableObject acForm, "C:\Export\Forms\frmCustomer.frm"

' Explicit object name (different from the file name)
ImportScriptableObject acForm, "C:\Export\Forms\frmCustomer.frm", "frmCustomerCopy"
```

**Parameters:**

| Parameter    | Type           | Optional | Description |
|--------------|----------------|----------|-------------|
| `ObjType`    | `AcObjectType` | No       | `acForm`, `acReport`, `acMacro`, or `acModule` — must match the type the file was exported as. |
| `FilePath`   | String         | No       | Full path to the exported file. |
| `ObjectName` | String         | Yes      | Name to give the object in the destination ADP. If omitted, the file name (without extension) is used. |

Return value: `Boolean` (`True` = imported successfully).

**Notes:**

- If an object with the target name already exists, `Application.LoadFromText`
  overwrites it without prompting.
- For bound forms/reports, the referenced view/table/stored procedure must
  exist with a matching name and columns in the destination ADP's SQL
  Server backend.
- The destination ADP's VBA project references (Tools → References) should
  match the source ADP's, so that any contained code compiles without
  errors.

## Requirements

- Microsoft Access 2003 (ADP project)
- No additional VBA project references required (standard libraries only)
