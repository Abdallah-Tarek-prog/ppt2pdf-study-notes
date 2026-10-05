// Drives the installed desktop PowerPoint to export each presentation.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Ppt2Pdf
{
    sealed class Converter
    {
        const int MsoTrue = -1;
        const int MsoFalse = 0;
        const int MsoPlaceholder = 14;
        const int PlaceholderBody = 2;

        readonly int mode;   // 1 slides, 2 slides with notes, 3 both
        dynamic app;
        bool quitWhenDone;
        public bool Aborted;
        public int Failures;

        public Converter(int mode) { this.mode = mode; }

        public void ConvertPath(string path)
        {
            if (Aborted) return;
            if (Directory.Exists(path))
            {
                string[] files;
                try
                {
                    files = Directory.GetFiles(path);
                }
                catch (Exception e)
                {
                    ConsoleIO.WriteLine("Could not read folder " + path + ": " + Describe(e));
                    Failures++;
                    return;
                }
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                foreach (string file in files)
                {
                    if (Aborted) return;
                    // "~$" files are PowerPoint's lock files for presentations that are open.
                    if (IsPresentation(file) && !Path.GetFileName(file).StartsWith("~$")) ConvertFile(Path.GetFullPath(file));
                }
            }
            else if (File.Exists(path))
            {
                if (IsPresentation(path)) ConvertFile(Path.GetFullPath(path));
                else ConsoleIO.WriteLine("Skipped (not a PPT or PPTX): " + path);
            }
            else
            {
                ConsoleIO.WriteLine("Not found: " + path);
            }
        }

        static bool IsPresentation(string path)
        {
            string extension = Path.GetExtension(path).ToLowerInvariant();
            return extension == ".ppt" || extension == ".pptx";
        }

        void ConvertFile(string source)
        {
            string folder = Path.GetDirectoryName(source);
            string stem = Path.GetFileNameWithoutExtension(source);
            string slidesPdf = Path.Combine(folder, stem + ".pdf");
            string notesPdf = mode == 2 ? slidesPdf : Path.Combine(folder, stem + " - Lec Notes.pdf");
            bool wantSlides = mode != 2 && !File.Exists(slidesPdf);
            bool wantNotes = mode != 1 && !File.Exists(notesPdf);
            if (!wantSlides && !wantNotes)
            {
                ConsoleIO.WriteLine("Skipped (output already exists): " + stem);
                return;
            }
            if (!StartPowerPoint()) return;

            dynamic presentation;
            try
            {
                presentation = app.Presentations.Open(source, MsoTrue, MsoFalse, MsoFalse);
            }
            catch (Exception e)
            {
                ConsoleIO.WriteLine("Failed to open " + source + ": " + Describe(e));
                Failures++;
                return;
            }

            try
            {
                bool slidesCreated = false;
                if (wantSlides)
                {
                    try
                    {
                        ExportPdf(presentation, slidesPdf);
                        ConsoleIO.WriteLine("Created: " + slidesPdf);
                        slidesCreated = true;
                    }
                    catch (Exception e)
                    {
                        ConsoleIO.WriteLine("Failed (slides): " + Describe(e));
                        Failures++;
                    }
                }
                if (wantNotes)
                {
                    try
                    {
                        CreateNotesPdf(presentation, notesPdf, slidesCreated ? slidesPdf : null);
                        ConsoleIO.WriteLine("Created: " + notesPdf);
                    }
                    catch (Exception e)
                    {
                        ConsoleIO.WriteLine("Failed (study PDF): " + Describe(e));
                        Failures++;
                    }
                }
            }
            finally
            {
                try { presentation.Close(); } catch (Exception) { }
            }
        }

        bool StartPowerPoint()
        {
            if (app != null) return true;
            try
            {
                bool alreadyRunning = Process.GetProcessesByName("POWERPNT").Length > 0;
                Type type = Type.GetTypeFromProgID("PowerPoint.Application");
                if (type == null) throw new InvalidOperationException("Microsoft PowerPoint is not installed.");
                app = Activator.CreateInstance(type);
                quitWhenDone = !alreadyRunning;
                return true;
            }
            catch (Exception e)
            {
                ConsoleIO.WriteLine("Could not start the installed PowerPoint: " + Describe(e));
                Failures++;
                Aborted = true;
                return false;
            }
        }

        // Closes PowerPoint only if this run started it and nothing else is open in it.
        public void Finish()
        {
            if (app == null) return;
            try
            {
                if (quitWhenDone && (int)app.Presentations.Count == 0) app.Quit();
            }
            catch (Exception) { }
            try { Marshal.FinalReleaseComObject((object)app); } catch (Exception) { }
            app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // Exports to a temporary name first so a failed export never leaves a
        // half-written PDF that later runs would skip as "already exists".
        static void ExportPdf(dynamic presentation, string path)
        {
            string temporary = Path.Combine(Path.GetDirectoryName(path), "~" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".pdf");
            try
            {
                // PDF, screen intent, no frames, slides only, hidden slides left out,
                // and an explicit empty print range (PowerPoint rejects the call without it).
                Call(presentation, "ExportAsFixedFormat", temporary, 2, 1, MsoFalse, 1, 1, MsoFalse, new DispatchWrapper(null));
                if (!File.Exists(temporary)) throw new IOException("PowerPoint did not create the PDF.");
                File.Move(temporary, path);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
            }
        }

        void CreateNotesPdf(dynamic presentation, string outputPath, string existingSlidesPdf)
        {
            string temp = Path.Combine(Path.GetTempPath(), "ppt-to-pdf-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(temp);
            try
            {
                string slidesPdf = existingSlidesPdf;
                if (slidesPdf == null)
                {
                    slidesPdf = Path.Combine(temp, "slides.pdf");
                    ExportPdf(presentation, slidesPdf);
                }

                // Hidden slides are not in the exported PDF, so skip their notes too.
                List<string> notes = new List<string>();
                List<int> slideNumbers = new List<int>();
                dynamic slides = presentation.Slides;
                int count = slides.Count;
                for (int n = 1; n <= count; n++)
                {
                    dynamic slide = slides.Item(n);
                    if ((int)slide.SlideShowTransition.Hidden == MsoTrue) continue;
                    notes.Add(NoteText.Strip(GetNotes(slide)));
                    slideNumbers.Add(n);
                }

                double slideWidth = presentation.PageSetup.SlideWidth;
                double slideHeight = presentation.PageSetup.SlideHeight;
                int pixelWidth = Math.Max(1, (int)Math.Ceiling(slideWidth * 0.5));
                int pixelHeight = Math.Max(1, (int)Math.Ceiling(slideHeight * 0.5));
                List<int[]> footers = new List<int[]>();
                for (int i = 0; i < notes.Count; i++)
                {
                    footers.Add(notes[i].Length == 0 ? null :
                        FooterColor(slides.Item(slideNumbers[i]), Path.Combine(temp, "slide.png"), pixelWidth, pixelHeight));
                }

                byte[] pdf = NotesPdf.Build(File.ReadAllBytes(slidesPdf), notes, footers);
                string temporaryOutput = outputPath + ".tmp";
                File.WriteAllBytes(temporaryOutput, pdf);
                if (File.Exists(outputPath))
                {
                    File.Delete(temporaryOutput);
                    throw new IOException("Output already exists: " + outputPath);
                }
                File.Move(temporaryOutput, outputPath);
            }
            finally
            {
                try { Directory.Delete(temp, true); } catch (Exception) { }
            }
        }

        static string GetNotes(dynamic slide)
        {
            dynamic shapes = slide.NotesPage.Shapes;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic shape = shapes.Item(i);
                if ((int)shape.Type == MsoPlaceholder && (int)shape.PlaceholderFormat.Type == PlaceholderBody)
                    return (string)shape.TextFrame.TextRange.Text ?? "";
            }
            return "";
        }

        // The most common colour along the bottom of the slide, rendered at half size.
        static int[] FooterColor(dynamic slide, string pngPath, int width, int height)
        {
            try
            {
                Call(slide, "Export", pngPath, "PNG", width, height);
                using (Bitmap image = new Bitmap(pngPath))
                {
                    Dictionary<int, int> counts = new Dictionary<int, int>();
                    Dictionary<int, long[]> totals = new Dictionary<int, long[]>();
                    List<int> firstSeen = new List<int>();
                    int w = image.Width, h = image.Height;
                    for (int y = (int)(h * 0.92); y < h; y += 2)
                    {
                        for (int x = (int)(w * 0.03); x < (int)(w * 0.97); x += 2)
                        {
                            Color pixel = image.GetPixel(x, y);
                            int r = Blend(pixel.R, pixel.A), g = Blend(pixel.G, pixel.A), b = Blend(pixel.B, pixel.A);
                            int group = (r / 16) << 8 | (g / 16) << 4 | (b / 16);
                            long[] total;
                            if (!totals.TryGetValue(group, out total))
                            {
                                total = new long[3];
                                totals[group] = total;
                                counts[group] = 0;
                                firstSeen.Add(group);
                            }
                            counts[group]++;
                            total[0] += r;
                            total[1] += g;
                            total[2] += b;
                        }
                    }
                    if (firstSeen.Count == 0) return new int[] { 255, 255, 255 };

                    int best = firstSeen[0];
                    foreach (int group in firstSeen)
                        if (counts[group] > counts[best]) best = group;
                    long[] sum = totals[best];
                    int samples = counts[best];
                    return new int[] {
                        (int)Math.Round((double)sum[0] / samples),
                        (int)Math.Round((double)sum[1] / samples),
                        (int)Math.Round((double)sum[2] / samples) };
                }
            }
            catch (Exception)
            {
                return new int[] { 255, 255, 255 };   // falls back to the neutral grey panel
            }
            finally
            {
                try { File.Delete(pngPath); } catch (Exception) { }
            }
        }

        // Plain IDispatch call; PowerPoint converts the arguments itself, which
        // is more forgiving than the C# dynamic binder for these methods.
        static object Call(object target, string method, params object[] arguments)
        {
            return target.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, target, arguments);
        }

        static int Blend(int channel, int alpha)
        {
            return (channel * alpha + 255 * (255 - alpha) + 127) / 255;
        }

        static string Describe(Exception e)
        {
            while (e.InnerException != null && (e is TargetInvocationException)) e = e.InnerException;
            return e.Message.Trim();
        }
    }
}
