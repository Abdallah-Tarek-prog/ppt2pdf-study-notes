// Builds the "slides with notes" PDF: each slide that has speaker notes gets
// a panel below it, sized to fit the text and coloured to match the slide's footer.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Ppt2Pdf
{
    static class NotesPdf
    {
        static readonly int[] TextColor = { 31, 41, 55 };

        // 'notes' has one entry per PDF page (already trimmed); 'footers' holds
        // the footer colour for each page with notes.
        public static byte[] Build(byte[] slidesPdf, IList<string> notes, IList<int[]> footers)
        {
            PdfDocument document = PdfDocument.Load(slidesPdf);
            if (document.Trailer.Get("Encrypt") != null) throw new PdfException("The exported PDF is encrypted.");
            List<PdfRef> pages = document.GetPages();
            if (pages.Count != notes.Count) throw new PdfException("The number of exported slides does not match the notes.");

            bool anyNotes = false;
            foreach (string text in notes) anyNotes |= text.Length > 0;
            if (!anyNotes) return slidesPdf;

            string fonts = Path.Combine(Environment.GetEnvironmentVariable("WINDIR") ?? @"C:\Windows", "Fonts");
            PdfFont regular = new PdfFont(TrueTypeFont.Load(Path.Combine(fonts, "arial.ttf")));
            PdfFont bold = new PdfFont(TrueTypeFont.Load(Path.Combine(fonts, "arialbd.ttf")));

            PdfRewriter rewriter = new PdfRewriter(document);
            PdfRef regularRef = rewriter.Reserve();
            PdfRef boldRef = rewriter.Reserve();
            for (int i = 0; i < pages.Count; i++)
            {
                if (notes[i].Length == 0) continue;
                int[] footer = footers[i] ?? new int[] { 255, 255, 255 };
                AddPanel(document, rewriter, pages[i], notes[i], footer, regular, regularRef, bold, boldRef);
            }
            rewriter.Replace(regularRef.Number, regular.CreateFontObject(rewriter));
            rewriter.Replace(boldRef.Number, bold.CreateFontObject(rewriter));

            // The slides' accessibility tags do not cover the notes, so drop them
            // rather than leave readers with a structure that skips the notes.
            PdfRef rootRef = document.Trailer.Get("Root") as PdfRef;
            PdfDict catalog = document.Resolve(rootRef) as PdfDict;
            if (rootRef != null && catalog != null)
            {
                catalog = catalog.Clone();
                catalog.Remove("StructTreeRoot");
                catalog.Remove("MarkInfo");
                rewriter.Replace(rootRef.Number, catalog);
            }
            return rewriter.Write(document.Trailer);
        }

        static void AddPanel(PdfDocument document, PdfRewriter rewriter, PdfRef pageRef, string notes, int[] footer,
                             PdfFont regular, PdfRef regularRef, PdfFont bold, PdfRef boldRef)
        {
            PdfDict page = document.Resolve(pageRef) as PdfDict;
            if (page == null) throw new PdfException("Missing page in the exported PDF.");
            double[] media = ReadBox(document, document.GetInherited(page, "MediaBox")) ?? new double[] { 0, 0, 612, 792 };
            double[] crop = ReadBox(document, document.GetInherited(page, "CropBox"));
            double[] box = media;
            if (crop != null)
            {
                box = new double[] {
                    Math.Max(media[0], crop[0]), Math.Max(media[1], crop[1]),
                    Math.Min(media[2], crop[2]), Math.Min(media[3], crop[3]) };
            }
            double width = box[2] - box[0];
            double height = box[3] - box[1];

            PdfDict resources = document.Resolve(document.GetInherited(page, "Resources")) as PdfDict;
            resources = resources != null ? resources.Clone() : new PdfDict();
            PdfDict fontResources = document.Resolve(resources.Get("Font")) as PdfDict;
            fontResources = fontResources != null ? fontResources.Clone() : new PdfDict();
            string regularName = UniqueName(fontResources, "NotesArial");
            string boldName = UniqueName(fontResources, "NotesArialBold");
            fontResources.Set(regularName, regularRef);
            fontResources.Set(boldName, boldRef);
            resources.Set("Font", fontResources);

            double panelHeight;
            string panel = DrawPanel(width, notes, footer, regular, regularName, bold, boldName, out panelHeight);

            // New page: the panel along the bottom, the slide moved up above it and
            // clipped to its own area. The panel is drawn first so the slide's
            // bottom edge blends cleanly over it.
            double dx = -box[0], dy = panelHeight - box[1];
            string before = string.Format("q\n{0}Q\nq 1 0 0 1 {1} {2} cm\n{3} {4} {5} {6} re W n\n",
                panel, N(dx), N(dy), N(box[0]), N(box[1]), N(width), N(height));
            string after = "Q\n";

            PdfArray contents = new PdfArray();
            contents.Items.Add(rewriter.Add(ContentStream(before)));
            PdfObject original = page.Get("Contents");
            PdfArray originalArray = document.Resolve(original) as PdfArray;
            if (originalArray != null) contents.Items.AddRange(originalArray.Items);
            else if (original != null) contents.Items.Add(original);
            contents.Items.Add(rewriter.Add(ContentStream(after)));

            PdfDict updated = page.Clone();
            updated.Set("MediaBox", Box(0, 0, width, height + panelHeight));
            updated.Remove("CropBox");
            updated.Remove("BleedBox");
            updated.Remove("TrimBox");
            updated.Remove("ArtBox");
            updated.Remove("Thumb");
            updated.Set("Resources", resources);
            updated.Set("Contents", contents);
            PdfArray annotations = document.Resolve(page.Get("Annots")) as PdfArray;
            if (annotations != null) updated.Set("Annots", MoveAnnotations(document, rewriter, annotations, dx, dy));
            rewriter.Replace(pageRef.Number, updated);
        }

        // Links sit in page coordinates, so they move up with the slide.
        static PdfArray MoveAnnotations(PdfDocument document, PdfRewriter rewriter, PdfArray annotations, double dx, double dy)
        {
            PdfArray moved = new PdfArray();
            foreach (PdfObject item in annotations.Items)
            {
                PdfDict annotation = document.Resolve(item) as PdfDict;
                if (annotation == null)
                {
                    moved.Items.Add(item);
                    continue;
                }
                annotation = annotation.Clone();
                foreach (string key in new string[] { "Rect", "QuadPoints" })
                {
                    PdfArray points = document.Resolve(annotation.Get(key)) as PdfArray;
                    if (points == null) continue;
                    PdfArray shifted = new PdfArray();
                    for (int i = 0; i < points.Items.Count; i++)
                    {
                        PdfNumber number = document.Resolve(points.Items[i]) as PdfNumber;
                        shifted.Items.Add(number == null ? points.Items[i] : new PdfNumber(number.Value + (i % 2 == 0 ? dx : dy)));
                    }
                    annotation.Set(key, shifted);
                }
                PdfRef reference = item as PdfRef;
                if (reference != null)
                {
                    rewriter.Replace(reference.Number, annotation);
                    moved.Items.Add(reference);
                }
                else
                {
                    moved.Items.Add(annotation);
                }
            }
            return moved;
        }

        // Panel layout, in points.
        static string DrawPanel(double pageWidth, string notes, int[] footer, PdfFont regular, string regularName,
                                PdfFont bold, string boldName, out double panelHeight)
        {
            double padding = Math.Max(24, pageWidth * 0.045);
            double bodySize = Math.Max(15, Math.Min(20, pageWidth * 0.024));
            double titleSize = bodySize * 1.3;
            double lineHeight = bodySize * 1.38;
            List<string> lines = Wrap(notes, regular, bodySize, pageWidth - 2 * padding);
            panelHeight = 2 * padding + titleSize + 16 + lines.Count * lineHeight;
            int[] background, accent;
            PanelColors(footer, out background, out accent);

            StringBuilder ops = new StringBuilder();
            ops.Append(Rgb(background)).Append(" rg\n");
            ops.AppendFormat("0 0 {0} {1} re f\n", N(pageWidth), N(panelHeight));
            ops.Append(Rgb(accent)).Append(" RG\n3 w\n");
            ops.AppendFormat("0 {0} m {1} {0} l S\n", N(panelHeight - 1.5), N(pageWidth));

            double titleY = panelHeight - padding - titleSize;
            ops.Append("BT\n");
            ops.AppendFormat("/{0} {1} Tf\n{2} rg\n", boldName, N(titleSize), Rgb(accent));
            ops.AppendFormat("1 0 0 1 {0} {1} Tm {2} Tj\n", N(padding), N(titleY), bold.Encode("Lec Notes"));

            ops.AppendFormat("/{0} {1} Tf\n{2} rg\n", regularName, N(bodySize), Rgb(TextColor));
            double firstY = titleY - 16 - bodySize;
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].Length == 0) continue;
                ops.AppendFormat("1 0 0 1 {0} {1} Tm {2} Tj\n", N(padding), N(firstY - i * lineHeight), regular.Encode(lines[i]));
            }
            ops.Append("ET\n");
            return ops.ToString();
        }

        public static List<string> Wrap(string text, PdfFont font, double size, double width)
        {
            List<string> lines = new List<string>();
            string normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
            foreach (string paragraph in normalized.Split('\n'))
            {
                List<string> words = NoteText.SplitWords(paragraph);
                if (words.Count == 0)
                {
                    lines.Add("");
                    continue;
                }

                string line = "";
                foreach (string word in words)
                {
                    string candidate = line.Length > 0 ? line + " " + word : word;
                    if (font.Width(candidate, size) <= width)
                    {
                        line = candidate;
                        continue;
                    }
                    if (line.Length > 0)
                    {
                        lines.Add(line);
                        line = "";
                    }
                    foreach (string letter in NoteText.CodePoints(word))
                    {
                        if (line.Length > 0 && font.Width(line + letter, size) > width)
                        {
                            lines.Add(line);
                            line = "";
                        }
                        line += letter;
                    }
                }
                if (line.Length > 0) lines.Add(line);
            }
            return lines;
        }

        public static void PanelColors(int[] footer, out int[] background, out int[] accent)
        {
            int min = Math.Min(footer[0], Math.Min(footer[1], footer[2]));
            int max = Math.Max(footer[0], Math.Max(footer[1], footer[2]));
            if (min > 235 && max - min < 20)
            {
                background = new int[] { 247, 249, 250 };
                accent = new int[] { 71, 85, 105 };
                return;
            }

            background = new int[3];
            for (int i = 0; i < 3; i++) background[i] = (int)Math.Round(0.18 * footer[i] + 0.82 * 255);
            accent = footer;
            foreach (double factor in new double[] { 1, 0.8, 0.65, 0.5, 0.4, 0.3 })
            {
                int[] candidate = new int[3];
                for (int i = 0; i < 3; i++) candidate[i] = (int)Math.Round(footer[i] * factor);
                if (Contrast(candidate, background) >= 4.5)
                {
                    accent = candidate;
                    break;
                }
            }
        }

        static double Contrast(int[] first, int[] second)
        {
            double a = Luminance(first), b = Luminance(second);
            double high = Math.Max(a, b), low = Math.Min(a, b);
            return (high + 0.05) / (low + 0.05);
        }

        static double Luminance(int[] color)
        {
            double[] values = new double[3];
            for (int i = 0; i < 3; i++)
            {
                double v = color[i] / 255.0;
                values[i] = v <= 0.04045 ? v / 12.92 : Math.Pow((v + 0.055) / 1.055, 2.4);
            }
            return 0.2126 * values[0] + 0.7152 * values[1] + 0.0722 * values[2];
        }

        static PdfStream ContentStream(string operators)
        {
            PdfDict dict = new PdfDict();
            dict.Set("Filter", new PdfName("FlateDecode"));
            return new PdfStream(dict, Zlib.Deflate(Encoding.ASCII.GetBytes(operators)));
        }

        static string UniqueName(PdfDict dict, string name)
        {
            string candidate = name;
            for (int i = 2; dict.Get(candidate) != null; i++) candidate = name + i.ToString(CultureInfo.InvariantCulture);
            return candidate;
        }

        static double[] ReadBox(PdfDocument document, PdfObject value)
        {
            PdfArray array = document.Resolve(value) as PdfArray;
            if (array == null || array.Items.Count != 4) return null;
            double[] numbers = new double[4];
            for (int i = 0; i < 4; i++)
            {
                PdfNumber number = document.Resolve(array.Items[i]) as PdfNumber;
                if (number == null) return null;
                numbers[i] = number.Value;
            }
            return new double[] {
                Math.Min(numbers[0], numbers[2]), Math.Min(numbers[1], numbers[3]),
                Math.Max(numbers[0], numbers[2]), Math.Max(numbers[1], numbers[3]) };
        }

        static PdfArray Box(double x0, double y0, double x1, double y1)
        {
            PdfArray array = new PdfArray();
            array.Items.Add(new PdfNumber(x0));
            array.Items.Add(new PdfNumber(y0));
            array.Items.Add(new PdfNumber(x1));
            array.Items.Add(new PdfNumber(y1));
            return array;
        }

        static string Rgb(int[] color)
        {
            return N(color[0] / 255.0) + " " + N(color[1] / 255.0) + " " + N(color[2] / 255.0);
        }

        static string N(double value)
        {
            return PdfNumber.Format(value);
        }
    }

    // An embedded, subsetted TrueType font (Type 0 / Identity-H, so any character works).
    sealed class PdfFont
    {
        readonly TrueTypeFont font;
        readonly SortedDictionary<int, int> used = new SortedDictionary<int, int>();   // glyph -> character

        public PdfFont(TrueTypeFont font) { this.font = font; }

        public double Width(string text, double size)
        {
            double total = 0;
            foreach (int codePoint in NoteText.CodePointValues(text)) total += font.Advance(font.GlyphFor(codePoint));
            return total * size / font.UnitsPerEm;
        }

        public string Encode(string text)
        {
            StringBuilder hex = new StringBuilder("<");
            foreach (int codePoint in NoteText.CodePointValues(text))
            {
                int glyph = font.GlyphFor(codePoint);
                if (!used.ContainsKey(glyph)) used[glyph] = glyph == 0 ? 0 : codePoint;
                hex.Append(glyph.ToString("X4"));
            }
            return hex.Append('>').ToString();
        }

        public PdfDict CreateFontObject(PdfRewriter rewriter)
        {
            string baseName = SubsetTag() + "+" + font.PostScriptName.Replace(" ", "");
            byte[] subset = font.Subset(used);

            PdfDict fileDict = new PdfDict();
            fileDict.Set("Filter", new PdfName("FlateDecode"));
            fileDict.Set("Length1", new PdfNumber(subset.Length));
            PdfRef fontFile = rewriter.Add(new PdfStream(fileDict, Zlib.Deflate(subset)));

            PdfDict descriptor = new PdfDict();
            descriptor.Set("Type", new PdfName("FontDescriptor"));
            descriptor.Set("FontName", new PdfName(baseName));
            descriptor.Set("Flags", new PdfNumber(32));
            PdfArray bbox = new PdfArray();
            foreach (int v in new int[] { font.XMin, font.YMin, font.XMax, font.YMax }) bbox.Items.Add(new PdfNumber(Scale(v)));
            descriptor.Set("FontBBox", bbox);
            descriptor.Set("ItalicAngle", new PdfNumber(0));
            descriptor.Set("Ascent", new PdfNumber(Scale(font.Ascender)));
            descriptor.Set("Descent", new PdfNumber(Scale(font.Descender)));
            descriptor.Set("CapHeight", new PdfNumber(Scale(font.CapHeight)));
            descriptor.Set("StemV", new PdfNumber(50 + (int)Math.Pow(font.WeightClass / 65.0, 2)));
            descriptor.Set("FontFile2", fontFile);

            PdfArray widths = new PdfArray();
            foreach (int glyph in used.Keys)
            {
                widths.Items.Add(new PdfNumber(glyph));
                PdfArray single = new PdfArray();
                single.Items.Add(new PdfNumber(Scale(font.Advance(glyph))));
                widths.Items.Add(single);
            }

            PdfDict systemInfo = new PdfDict();
            systemInfo.Set("Registry", new PdfString(Encoding.ASCII.GetBytes("(Adobe)")));
            systemInfo.Set("Ordering", new PdfString(Encoding.ASCII.GetBytes("(Identity)")));
            systemInfo.Set("Supplement", new PdfNumber(0));

            PdfDict cidFont = new PdfDict();
            cidFont.Set("Type", new PdfName("Font"));
            cidFont.Set("Subtype", new PdfName("CIDFontType2"));
            cidFont.Set("BaseFont", new PdfName(baseName));
            cidFont.Set("CIDSystemInfo", systemInfo);
            cidFont.Set("FontDescriptor", rewriter.Add(descriptor));
            cidFont.Set("DW", new PdfNumber(Scale(font.Advance(0))));
            cidFont.Set("W", widths);
            cidFont.Set("CIDToGIDMap", new PdfName("Identity"));

            PdfArray descendants = new PdfArray();
            descendants.Items.Add(rewriter.Add(cidFont));

            PdfDict toUnicodeDict = new PdfDict();
            toUnicodeDict.Set("Filter", new PdfName("FlateDecode"));

            PdfDict type0 = new PdfDict();
            type0.Set("Type", new PdfName("Font"));
            type0.Set("Subtype", new PdfName("Type0"));
            type0.Set("BaseFont", new PdfName(baseName));
            type0.Set("Encoding", new PdfName("Identity-H"));
            type0.Set("DescendantFonts", descendants);
            type0.Set("ToUnicode", rewriter.Add(new PdfStream(toUnicodeDict, Zlib.Deflate(Encoding.ASCII.GetBytes(ToUnicodeMap())))));
            return type0;
        }

        int Scale(int value)
        {
            return (int)Math.Round(value * 1000.0 / font.UnitsPerEm);
        }

        string SubsetTag()
        {
            StringBuilder key = new StringBuilder(font.PostScriptName);
            foreach (int glyph in used.Keys) key.Append(',').Append(glyph);
            byte[] hash;
            using (MD5 md5 = MD5.Create()) hash = md5.ComputeHash(Encoding.UTF8.GetBytes(key.ToString()));
            StringBuilder tag = new StringBuilder();
            for (int i = 0; i < 6; i++) tag.Append((char)('A' + hash[i] % 26));
            return tag.ToString();
        }

        string ToUnicodeMap()
        {
            List<string> entries = new List<string>();
            foreach (KeyValuePair<int, int> pair in used)
            {
                if (pair.Value == 0 || (pair.Value >= 0xD800 && pair.Value <= 0xDFFF)) continue;
                StringBuilder utf16 = new StringBuilder();
                foreach (char c in char.ConvertFromUtf32(pair.Value)) utf16.Append(((int)c).ToString("X4"));
                entries.Add("<" + pair.Key.ToString("X4") + "> <" + utf16 + ">");
            }

            StringBuilder map = new StringBuilder();
            map.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n");
            map.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            map.Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n");
            map.Append("1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");
            for (int start = 0; start < entries.Count; start += 100)
            {
                int count = Math.Min(100, entries.Count - start);
                map.Append(count).Append(" beginbfchar\n");
                for (int i = start; i < start + count; i++) map.Append(entries[i]).Append('\n');
                map.Append("endbfchar\n");
            }
            map.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
            return map.ToString();
        }
    }

    // Text helpers for wrapping notes: what counts as a space, and splitting by character.
    static class NoteText
    {
        public static bool IsSpace(int c)
        {
            return (c >= 0x09 && c <= 0x0D) || (c >= 0x1C && c <= 0x20) || c == 0x85 || c == 0xA0 ||
                   c == 0x1680 || (c >= 0x2000 && c <= 0x200A) || c == 0x2028 || c == 0x2029 ||
                   c == 0x202F || c == 0x205F || c == 0x3000;
        }

        public static string Strip(string text)
        {
            int start = 0, end = text.Length;
            while (start < end && IsSpace(text[start])) start++;
            while (end > start && IsSpace(text[end - 1])) end--;
            return text.Substring(start, end - start);
        }

        public static List<string> SplitWords(string text)
        {
            List<string> words = new List<string>();
            StringBuilder word = new StringBuilder();
            foreach (char c in text)
            {
                if (IsSpace(c))
                {
                    if (word.Length > 0) words.Add(word.ToString());
                    word.Length = 0;
                }
                else
                {
                    word.Append(c);
                }
            }
            if (word.Length > 0) words.Add(word.ToString());
            return words;
        }

        public static IEnumerable<string> CodePoints(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    yield return text.Substring(i, 2);
                    i++;
                }
                else
                {
                    yield return text.Substring(i, 1);
                }
            }
        }

        public static IEnumerable<int> CodePointValues(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    yield return char.ConvertToUtf32(text[i], text[i + 1]);
                    i++;
                }
                else
                {
                    yield return text[i];
                }
            }
        }
    }
}
