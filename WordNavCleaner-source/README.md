# docx-heading-cleaner

`docx-heading-cleaner` removes Word navigation outline levels from `.docx`
paragraphs while preserving the original heading styles for later manual use.

First version supports only ALL mode:

```bash
./docx-heading-cleaner input.docx --mode all
./docx-heading-cleaner input.docx --mode all --out output.docx
```

The source file is never overwritten. By default, output is written as:

```text
input.cleaned.docx
```

What ALL mode does:

- paragraphs using navigation styles are moved to cloned `No Nav` styles;
- direct paragraph outline levels are removed;
- original `Heading 1` / `标题 1` styles stay untouched, so real headings can be
  manually restored in Word by applying the original heading style again.

The tool uses only the Python standard library.

## Windows drag-and-drop app

The GUI version is designed for Windows users who do not want to use a command
line. It opens a small window named `Word 导航标题清理器`; drag one or more
`.docx` files into the window, then click `开始清理`.

Outputs are written next to the original files:

```text
input.cleaned.docx
```

If the cleaned file already exists, the app asks whether to overwrite it.

Run the GUI during development:

```bash
pip install -r requirements.txt
python gui_app.py
```

Build the Windows EXE on a Windows machine:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows.ps1
```

The packaged app will be created at:

```text
dist\WordNavCleaner.exe
```

Notes:

- The EXE should be built on Windows. Building it on macOS will not produce a
  reliable Windows executable.
- The first version supports `.docx` only, not legacy `.doc` files.
- Because the app is not code-signed, Windows SmartScreen may show an unknown
  publisher warning.

## Build without installing Python locally

If you want a ready-to-send Windows app without setting up Python on your own
computer, use the included GitHub Actions workflow.

One-time setup:

1. Create a GitHub repository.
2. Upload these project files to the repository.
3. Open the repository page in GitHub.
4. Go to `Actions`.
5. Choose `Build Windows app`.
6. Click `Run workflow`.

After it finishes, open the completed workflow run and download the artifacts:

```text
WordNavCleaner-portable
WordNavCleaner-installer
```

Inside those downloads:

```text
WordNavCleaner.exe
WordNavCleanerSetup.exe
```

Recommended sharing choice:

- Send `WordNavCleanerSetup.exe` to users who want a normal installer.
- Send `WordNavCleaner.exe` to users who want a portable app they can double
  click directly.

The GitHub build uses a Windows cloud runner, installs the build tools there,
runs the tests, builds the portable EXE, then builds the installer.
