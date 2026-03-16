# aero-msstyles DarkMode Class Parser

A lightweight Windows utility that parses and displays **DarkMode class mappings** embedded in `.msstyles` theme files (e.g. `aero.msstyles`).

It reads the `CMap` resource section of a Visual Styles file and extracts all `DarkMode::*` class variants, displaying them in a hierarchical tree view.

---

## Features

- 📂 Open any `.msstyles` file via a file dialog
- 🌲 Displays DarkMode class groups as a tree (base class → variants)
- 💾 Persists the last opened file path via the Windows Registry
- 🔍 Enumerates all `CMap` resources (not just ID=1)
- ✅ Supports MSVC and MinGW/GCC builds

---

## Screenshot

> Tree view showing DarkMode base classes and their variants parsed from `aero.msstyles`.

---

## Building

### Requirements

- Windows SDK
- A C++ compiler (MSVC or MinGW-w64)
- Links against: `Comctl32.lib`, `Shlwapi.lib`, `User32.lib`, `Advapi32.lib`, `Comdlg32.lib`, `Gdi32.lib`

### MSVC (Visual Studio)

Open a **Developer Command Prompt** and run:

```bat
cl /W4 /EHsc main.cpp /link Comctl32.lib Shlwapi.lib User32.lib Advapi32.lib Comdlg32.lib Gdi32.lib /SUBSYSTEM:WINDOWS
```

### MinGW / GCC

```bat
g++ -municode -mwindows main.cpp -o aero-darkmode-parser.exe -lcomctl32 -lshlwapi -luser32 -ladvapi32 -lcomdlg32 -lgdi32
```

---

## Usage

1. Run `aero-darkmode-parser.exe`
2. A file picker will open on first launch — select a `.msstyles` file  
   (typically at `C:\Windows\Resources\Themes\aero\aero.msstyles`)
3. The tree populates with all discovered `DarkMode` class groups
4. Use **File → Open .msstyles...** to load a different file

The last used path is remembered in the registry under:
```
HKCU\Software\Microsoft\Windows\CurrentVersion\ThemeManager
```

---

## How It Works

`.msstyles` files are PE-format DLLs containing embedded resources. Among these is the `CMap` resource type, which holds null-delimited wide-string entries mapping theme class names.

This tool:
1. Loads the `.msstyles` as a data-only module (`LOAD_LIBRARY_AS_DATAFILE`)
2. Enumerates all `CMap` resources via `EnumResourceNames`
3. Scans each blob for entries containing `DarkMode`
4. Groups them by base class name (the part after `::`)
5. Displays the hierarchy in a TreeView control

---

## Project Structure

```
aero-darkmode-parser/
├── main.cpp          # Full source (single-file Win32 application)
└── README.md
```

---

## License

MIT License. See [LICENSE](LICENSE).
