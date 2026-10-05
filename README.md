# ppt2pdf

Convert multiple `.ppt`/`.pptx` lectures to PDFs on Windows. Choose slides, slides with speaker notes below them, or both. Notes panels grow to fit the text and match each slide's footer colour.

It is one small program, `ppt2pdf.exe`, with nothing to install besides PowerPoint.

## Quick setup

1. Download **`ppt2pdf.exe`** from [Releases](../../releases/latest). If your browser blocks `.exe` downloads, take `ppt2pdf-windows.zip` instead and extract it.
2. Make sure desktop **Microsoft PowerPoint** is installed and activated. The converter uses it in the background; PowerPoint Online is not supported.
3. Double-click **`ppt2pdf.exe`**, drag one or more presentations (or a folder) into `Files or folder:`, press **Enter**, then press `1`, `2`, or `3`.

You can also drag presentation files or a folder straight onto `ppt2pdf.exe`.

The EXE is not code-signed, so the first time Windows may show "Windows protected your PC". Click **More info**, then **Run anyway**.

### Terminal command

```powershell
.\ppt2pdf.exe                          # asks for files, then the format
.\ppt2pdf.exe run                      # same
.\ppt2pdf.exe "Lecture 1.pptx" Week2   # convert these files/folders (asks for the format)
.\ppt2pdf.exe --help
```

To use **`ppt2pdf run` from any folder**, put `ppt2pdf.exe` in a folder of its own, add that folder to your user `PATH`, then open a new terminal.

PDFs are saved beside each presentation:

| Choice | Output for `Lecture.pptx` |
| --- | --- |
| `1` Slides | `Lecture.pdf` |
| `2` Slides with notes | `Lecture.pdf` |
| `3` Both | `Lecture.pdf` and `Lecture - Lec Notes.pdf` |

Existing PDFs are skipped, never overwritten. If you switch between choices `1` and `2`, rename or move the existing `Lecture.pdf` first.

Folders are not searched recursively. Hidden slides are left out, together with their notes.

## Requirements

- Windows 10 or 11 (it uses the .NET Framework that is part of Windows).
- Desktop Microsoft PowerPoint. If PowerPoint is already open, ppt2pdf works alongside it and leaves your open presentations alone.

## Known limitation

Right-to-left notes text (Arabic, Hebrew) is drawn letter by letter: letters are not joined and the word order is not reversed. Slides themselves are unaffected, since PowerPoint draws them.

## Source code

All code is in [`src/`](src), in C#. To build `ppt2pdf.exe`, run:

```powershell
.\build.cmd
```

It uses the C# compiler that ships with Windows, so nothing needs to be installed.

How it works: PowerPoint exports the slides to PDF. For the notes version, ppt2pdf reads each slide's speaker notes, then rebuilds the PDF itself. Each page gets a notes panel underneath, with the text set in Arial from the Windows fonts folder.

## License

MIT licensed. See [LICENSE](LICENSE).
