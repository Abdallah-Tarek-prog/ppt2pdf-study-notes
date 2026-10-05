// ppt2pdf: convert PowerPoint presentations to PDF, optionally with the
// speaker notes printed below each slide.
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

[assembly: AssemblyTitle("ppt2pdf")]
[assembly: AssemblyProduct("ppt2pdf")]
[assembly: AssemblyDescription("Convert PowerPoint presentations to PDF, with optional speaker notes")]
[assembly: AssemblyCopyright("Copyright (c) 2026 Abdallah-Tarek-prog, MIT License")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace Ppt2Pdf
{
    static class Program
    {
        const string Version = "1.0.0";

        [STAThread]
        static int Main(string[] args)
        {
            bool ownWindow = ConsoleIO.OwnsWindow();
            if (ownWindow)
            {
                try { Console.Title = "PowerPoint to PDF Converter"; } catch (Exception) { }
            }

            int exitCode;
            try
            {
                exitCode = Run(args);
            }
            catch (Exception e)
            {
                ConsoleIO.WriteLine("Unexpected error: " + e.Message);
                exitCode = 1;
            }

            if (exitCode != 2 && ownWindow && !Console.IsInputRedirected)
            {
                ConsoleIO.WriteLine("");
                ConsoleIO.Write("Press any key to continue . . . ");
                Console.ReadKey(true);
                ConsoleIO.WriteLine("");
            }
            return exitCode == 2 ? 0 : exitCode;
        }

        // Returns 0 on success, 1 if anything failed, 2 after printing help.
        static int Run(string[] args)
        {
            List<string> paths = new List<string>();
            foreach (string arg in args)
            {
                // A quoted folder ending in a backslash ("C:\Lectures\") arrives with a stray quote.
                string cleaned = arg.Trim().TrimEnd('"');
                if (cleaned.Length > 0) paths.Add(cleaned);
            }

            if (paths.Count > 0)
            {
                string first = paths[0].ToLowerInvariant();
                if (first == "-h" || first == "--help" || first == "/?" || first == "help")
                {
                    PrintHelp();
                    return 2;
                }
                if (first == "--version" || first == "-v")
                {
                    ConsoleIO.WriteLine("ppt2pdf " + Version);
                    return 2;
                }
                if (first == "run") paths.RemoveAt(0);
            }

            if (paths.Count == 0)
            {
                ConsoleIO.WriteLine("Drag one or more .ppt/.pptx files, or a folder, into this window.");
                ConsoleIO.WriteLine("Then press Enter.");
                ConsoleIO.Write("Files or folder: ");
                string line = ConsoleIO.ReadLine();
                paths = ParsePaths(line ?? "");
                if (paths.Count == 0) return 0;
            }

            ConsoleIO.WriteLine("Choose the PDF format:");
            ConsoleIO.WriteLine("  1  Slides only");
            ConsoleIO.WriteLine("  2  Slides with Lec Notes below each slide");
            ConsoleIO.WriteLine("  3  Both");
            int mode = ConsoleIO.Choose("Enter 1, 2, or 3: ", "123");
            if (mode == 0) return 1;

            Converter converter = new Converter(mode);
            try
            {
                foreach (string path in paths) converter.ConvertPath(path);
            }
            finally
            {
                converter.Finish();
            }
            return converter.Failures > 0 ? 1 : 0;
        }

        // Text dropped into the console: one path (spaces allowed) or several, quoted as needed.
        static List<string> ParsePaths(string line)
        {
            List<string> paths = new List<string>();
            string trimmed = line.Trim();
            if (trimmed.Length == 0) return paths;
            if (File.Exists(trimmed) || Directory.Exists(trimmed))
            {
                paths.Add(trimmed);
                return paths;
            }

            StringBuilder current = new StringBuilder();
            bool quoted = false, hasToken = false;
            foreach (char c in trimmed)
            {
                if (c == '"')
                {
                    quoted = !quoted;
                    hasToken = true;
                }
                else if (char.IsWhiteSpace(c) && !quoted)
                {
                    if (hasToken) paths.Add(current.ToString());
                    current.Length = 0;
                    hasToken = false;
                }
                else
                {
                    current.Append(c);
                    hasToken = true;
                }
            }
            if (hasToken) paths.Add(current.ToString());
            paths.RemoveAll(delegate (string p) { return p.Length == 0; });
            return paths;
        }

        static void PrintHelp()
        {
            ConsoleIO.WriteLine("ppt2pdf " + Version + " - convert PowerPoint presentations to PDF");
            ConsoleIO.WriteLine("");
            ConsoleIO.WriteLine("Usage:");
            ConsoleIO.WriteLine("  ppt2pdf                      Ask for files or a folder, then the format");
            ConsoleIO.WriteLine("  ppt2pdf run                  Same as above");
            ConsoleIO.WriteLine("  ppt2pdf <files or folders>   Convert these (asks for the format)");
            ConsoleIO.WriteLine("");
            ConsoleIO.WriteLine("Formats:");
            ConsoleIO.WriteLine("  1  Slides only                 -> Lecture.pdf");
            ConsoleIO.WriteLine("  2  Slides with notes below     -> Lecture.pdf");
            ConsoleIO.WriteLine("  3  Both                        -> Lecture.pdf and Lecture - Lec Notes.pdf");
            ConsoleIO.WriteLine("");
            ConsoleIO.WriteLine("PDFs are saved next to each presentation. Existing PDFs are never overwritten.");
            ConsoleIO.WriteLine("Needs the desktop version of Microsoft PowerPoint.");
        }
    }

    // Console input and output that keep non-English file names intact.
    static class ConsoleIO
    {
        [DllImport("kernel32.dll")] static extern IntPtr GetStdHandle(int handle);
        [DllImport("kernel32.dll")] static extern bool GetConsoleMode(IntPtr handle, out uint mode);
        [DllImport("kernel32.dll")] static extern uint GetConsoleProcessList(uint[] processes, uint count);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern bool WriteConsoleW(IntPtr handle, string text, uint length, out uint written, IntPtr reserved);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern bool ReadConsoleW(IntPtr handle, [Out] char[] buffer, uint length, out uint read, IntPtr control);

        static readonly IntPtr Output = GetStdHandle(-11);
        static readonly IntPtr Input = GetStdHandle(-10);
        static readonly bool OutputIsConsole = IsConsole(Output);
        static readonly bool InputIsConsole = IsConsole(Input);
        static TextWriter redirectedOutput;
        static TextReader redirectedInput;

        static bool IsConsole(IntPtr handle)
        {
            uint mode;
            return handle != IntPtr.Zero && handle != new IntPtr(-1) && GetConsoleMode(handle, out mode);
        }

        // True when this window was opened just for ppt2pdf (double-click or drag
        // and drop), so it should wait for a key before closing.
        public static bool OwnsWindow()
        {
            try
            {
                return GetConsoleProcessList(new uint[4], 4) == 1;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static void Write(string text)
        {
            if (OutputIsConsole)
            {
                uint written;
                WriteConsoleW(Output, text, (uint)text.Length, out written, IntPtr.Zero);
                return;
            }
            if (redirectedOutput == null)
            {
                StreamWriter writer = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false));
                writer.AutoFlush = true;
                redirectedOutput = writer;
            }
            redirectedOutput.Write(text);
        }

        public static void WriteLine(string text)
        {
            Write(text + "\r\n");
        }

        public static string ReadLine()
        {
            if (!InputIsConsole)
            {
                if (redirectedInput == null) redirectedInput = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
                return redirectedInput.ReadLine();
            }
            StringBuilder line = new StringBuilder();
            char[] buffer = new char[4096];
            while (true)
            {
                uint read;
                if (!ReadConsoleW(Input, buffer, (uint)buffer.Length, out read, IntPtr.Zero) || read == 0)
                    return line.Length > 0 ? line.ToString() : null;
                line.Append(buffer, 0, (int)read);
                string text = line.ToString();
                int end = text.IndexOf('\n');
                if (end >= 0) return text.Substring(0, end).TrimEnd('\r');
            }
        }

        // Waits for one of the allowed keys, like the Windows "choice" command.
        // Returns its position (1-based), or 0 if input ended.
        public static int Choose(string prompt, string keys)
        {
            Write(prompt);
            if (!InputIsConsole)
            {
                while (true)
                {
                    string line = ReadLine();
                    if (line == null) return 0;
                    int at = line.Trim().Length > 0 ? keys.IndexOf(line.Trim()[0]) : -1;
                    if (at >= 0)
                    {
                        WriteLine(keys[at].ToString());
                        return at + 1;
                    }
                }
            }
            while (true)
            {
                ConsoleKeyInfo key = Console.ReadKey(true);
                int at = keys.IndexOf(key.KeyChar);
                if (at >= 0)
                {
                    WriteLine(keys[at].ToString());
                    return at + 1;
                }
                Write("\a");
            }
        }
    }
}
