// Reads a TrueType font (Arial from the Windows fonts folder) for text
// measurement, and writes a subset containing only the glyphs used.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Ppt2Pdf
{
    sealed class TrueTypeFont
    {
        readonly byte[] data;
        readonly Dictionary<string, int[]> tables = new Dictionary<string, int[]>();
        readonly Dictionary<int, int> cmap = new Dictionary<int, int>();
        readonly int numberOfHMetrics;
        readonly bool longLoca;

        public readonly string PostScriptName;
        public readonly int UnitsPerEm;
        public readonly int NumGlyphs;
        public readonly int XMin, YMin, XMax, YMax;
        public readonly int Ascender, Descender, CapHeight, WeightClass;

        public static TrueTypeFont Load(string path)
        {
            if (!File.Exists(path)) throw new FileNotFoundException("Font not found: " + path);
            return new TrueTypeFont(File.ReadAllBytes(path));
        }

        TrueTypeFont(byte[] data)
        {
            this.data = data;
            int numTables = U16(4);
            for (int i = 0; i < numTables; i++)
            {
                int record = 12 + 16 * i;
                string tag = Encoding.ASCII.GetString(data, record, 4);
                tables[tag] = new int[] { (int)U32(record + 8), (int)U32(record + 12) };
            }

            int head = Table("head");
            UnitsPerEm = U16(head + 18);
            XMin = S16(head + 36);
            YMin = S16(head + 38);
            XMax = S16(head + 40);
            YMax = S16(head + 42);
            longLoca = S16(head + 50) == 1;
            NumGlyphs = U16(Table("maxp") + 4);
            int hhea = Table("hhea");
            Ascender = S16(hhea + 4);
            Descender = S16(hhea + 6);
            numberOfHMetrics = U16(hhea + 34);

            CapHeight = Ascender * 7 / 10;
            WeightClass = 400;
            if (tables.ContainsKey("OS/2"))
            {
                int os2 = Table("OS/2");
                WeightClass = U16(os2 + 4);
                if (U16(os2) >= 2 && tables["OS/2"][1] >= 90) CapHeight = S16(os2 + 88);
            }

            ReadCmap();
            PostScriptName = ReadPostScriptName() ?? "Arial";
        }

        public int GlyphFor(int codePoint)
        {
            int glyph;
            return cmap.TryGetValue(codePoint, out glyph) ? glyph : 0;
        }

        public int Advance(int glyph)
        {
            int index = Math.Min(glyph, numberOfHMetrics - 1);
            return U16(Table("hmtx") + 4 * index);
        }

        int Table(string tag)
        {
            int[] entry;
            if (!tables.TryGetValue(tag, out entry)) throw new InvalidDataException("Font is missing the '" + tag + "' table.");
            return entry[0];
        }

        void ReadCmap()
        {
            int cmapTable = Table("cmap");
            int count = U16(cmapTable + 2);
            int best = -1, bestScore = 0;
            for (int i = 0; i < count; i++)
            {
                int record = cmapTable + 4 + 8 * i;
                int platform = U16(record), encoding = U16(record + 2);
                int subtable = cmapTable + (int)U32(record + 4);
                int format = U16(subtable);
                int score = 0;
                if (format == 12 && (platform == 3 && encoding == 10 || platform == 0)) score = 4;
                else if (format == 4 && platform == 3 && encoding == 1) score = 3;
                else if (format == 4 && platform == 0) score = 2;
                else if (format == 4 && platform == 3 && encoding == 0) score = 1;
                if (score > bestScore)
                {
                    bestScore = score;
                    best = subtable;
                }
            }
            if (best < 0) throw new InvalidDataException("Font has no usable character map.");

            if (U16(best) == 12)
            {
                long groups = U32(best + 12);
                for (long g = 0; g < groups; g++)
                {
                    int group = best + 16 + (int)(12 * g);
                    long start = U32(group), end = U32(group + 4), glyph = U32(group + 8);
                    for (long c = start; c <= end && c <= 0x10FFFF; c++) cmap[(int)c] = (int)(glyph + c - start);
                }
                return;
            }

            int segments = U16(best + 6) / 2;
            int ends = best + 14;
            int starts = ends + 2 * segments + 2;
            int deltas = starts + 2 * segments;
            int rangeOffsets = deltas + 2 * segments;
            for (int s = 0; s < segments; s++)
            {
                int end = U16(ends + 2 * s), start = U16(starts + 2 * s);
                int delta = S16(deltas + 2 * s), rangeOffset = U16(rangeOffsets + 2 * s);
                for (int c = start; c <= end && c != 0xFFFF; c++)
                {
                    int glyph;
                    if (rangeOffset == 0)
                    {
                        glyph = (c + delta) & 0xFFFF;
                    }
                    else
                    {
                        int address = rangeOffsets + 2 * s + rangeOffset + 2 * (c - start);
                        glyph = address + 1 < data.Length ? U16(address) : 0;
                        if (glyph != 0) glyph = (glyph + delta) & 0xFFFF;
                    }
                    if (glyph != 0 && glyph < NumGlyphs) cmap[c] = glyph;
                }
            }
        }

        string ReadPostScriptName()
        {
            if (!tables.ContainsKey("name")) return null;
            int name = Table("name");
            int count = U16(name + 2);
            int strings = name + U16(name + 4);
            string fallback = null;
            for (int i = 0; i < count; i++)
            {
                int record = name + 6 + 12 * i;
                int platform = U16(record), nameId = U16(record + 6);
                int length = U16(record + 8), offset = U16(record + 10);
                if (nameId != 6) continue;
                if (platform == 3) return Encoding.BigEndianUnicode.GetString(data, strings + offset, length);
                if (platform == 1) fallback = Encoding.ASCII.GetString(data, strings + offset, length);
            }
            return fallback;
        }

        // Builds a smaller font that keeps glyph numbers but empties every
        // glyph not listed. 'characters' maps used glyphs to a character.
        public byte[] Subset(IDictionary<int, int> characters)
        {
            HashSet<int> keep = new HashSet<int>(characters.Keys);
            keep.Add(0);
            Stack<int> pending = new Stack<int>(keep);
            while (pending.Count > 0)
            {
                foreach (int component in Components(pending.Pop()))
                    if (component < NumGlyphs && keep.Add(component)) pending.Push(component);
            }

            int glyf = Table("glyf");
            MemoryStream newGlyf = new MemoryStream();
            byte[] newLoca = new byte[4 * (NumGlyphs + 1)];
            for (int g = 0; g < NumGlyphs; g++)
            {
                PutU32(newLoca, 4 * g, (uint)newGlyf.Length);
                if (!keep.Contains(g)) continue;
                int start, end;
                GlyphRange(g, out start, out end);
                newGlyf.Write(data, glyf + start, end - start);
                while (newGlyf.Length % 4 != 0) newGlyf.WriteByte(0);
            }
            PutU32(newLoca, 4 * NumGlyphs, (uint)newGlyf.Length);

            byte[] hmtx = Copy("hmtx");
            for (int g = 0; g < NumGlyphs; g++)
            {
                if (keep.Contains(g)) continue;
                int at = g < numberOfHMetrics ? 4 * g : 4 * numberOfHMetrics + 2 * (g - numberOfHMetrics);
                int size = g < numberOfHMetrics ? 4 : 2;
                if (at + size <= hmtx.Length) Array.Clear(hmtx, at, size);
            }

            byte[] head = Copy("head");
            PutU32(head, 8, 0);          // checkSumAdjustment, fixed below
            head[50] = 0;
            head[51] = 1;                // long loca offsets

            SortedDictionary<string, byte[]> output = new SortedDictionary<string, byte[]>(StringComparer.Ordinal);
            output["head"] = head;
            output["hhea"] = Copy("hhea");
            output["maxp"] = Copy("maxp");
            output["hmtx"] = hmtx;
            output["loca"] = newLoca;
            output["glyf"] = newGlyf.ToArray();
            output["cmap"] = BuildCmap(characters);
            output["post"] = BuildPost();
            foreach (string tag in new string[] { "cvt ", "fpgm", "prep", "OS/2", "gasp" })
                if (tables.ContainsKey(tag)) output[tag] = Copy(tag);
            byte[] name = BuildName();
            if (name != null) output["name"] = name;

            byte[] font = Assemble(output);
            uint adjustment = unchecked(0xB1B0AFBA - Checksum(font, 0, font.Length));
            int headOffset = FindTableOffset(font, "head");
            PutU32(font, headOffset + 8, adjustment);
            return font;
        }

        IEnumerable<int> Components(int glyph)
        {
            int start, end;
            GlyphRange(glyph, out start, out end);
            if (end - start < 10) yield break;
            int glyf = Table("glyf") + start;
            if (S16(glyf) >= 0) yield break;
            int p = glyf + 10;
            while (true)
            {
                int flags = U16(p);
                yield return U16(p + 2);
                p += 4;
                p += (flags & 0x0001) != 0 ? 4 : 2;
                if ((flags & 0x0008) != 0) p += 2;
                else if ((flags & 0x0040) != 0) p += 4;
                else if ((flags & 0x0080) != 0) p += 8;
                if ((flags & 0x0020) == 0) yield break;
            }
        }

        void GlyphRange(int glyph, out int start, out int end)
        {
            int loca = Table("loca");
            if (longLoca)
            {
                start = (int)U32(loca + 4 * glyph);
                end = (int)U32(loca + 4 * glyph + 4);
            }
            else
            {
                start = 2 * U16(loca + 2 * glyph);
                end = 2 * U16(loca + 2 * glyph + 2);
            }
            if (end < start) end = start;
        }

        byte[] BuildCmap(IDictionary<int, int> characters)
        {
            SortedDictionary<int, int> map = new SortedDictionary<int, int>();
            foreach (KeyValuePair<int, int> pair in characters)
                if (pair.Value > 0 && pair.Value < 0xFFFF && !map.ContainsKey(pair.Value)) map[pair.Value] = pair.Key;

            int segments = map.Count + 1;
            int length = 16 + 8 * segments;
            byte[] table = new byte[12 + length];
            PutU16(table, 0, 0);
            PutU16(table, 2, 1);
            PutU16(table, 4, 3);
            PutU16(table, 6, 1);
            PutU32(table, 8, 12);
            int sub = 12;
            int power = 1, log = 0;
            while (power * 2 <= segments) { power *= 2; log++; }
            PutU16(table, sub, 4);
            PutU16(table, sub + 2, length);
            PutU16(table, sub + 6, 2 * segments);
            PutU16(table, sub + 8, 2 * power);
            PutU16(table, sub + 10, log);
            PutU16(table, sub + 12, 2 * segments - 2 * power);
            int ends = sub + 14, starts = ends + 2 * segments + 2, deltas = starts + 2 * segments;
            int i = 0;
            foreach (KeyValuePair<int, int> pair in map)
            {
                PutU16(table, ends + 2 * i, pair.Key);
                PutU16(table, starts + 2 * i, pair.Key);
                PutU16(table, deltas + 2 * i, (pair.Value - pair.Key) & 0xFFFF);
                i++;
            }
            PutU16(table, ends + 2 * i, 0xFFFF);
            PutU16(table, starts + 2 * i, 0xFFFF);
            PutU16(table, deltas + 2 * i, 1);
            return table;
        }

        byte[] BuildPost()
        {
            byte[] post = new byte[32];
            if (tables.ContainsKey("post") && tables["post"][1] >= 16)
                Buffer.BlockCopy(data, Table("post"), post, 0, 16);
            PutU32(post, 0, 0x00030000);
            return post;
        }

        byte[] BuildName()
        {
            if (!tables.ContainsKey("name")) return null;
            int name = Table("name");
            int count = U16(name + 2);
            int strings = name + U16(name + 4);
            List<int[]> records = new List<int[]>();
            for (int i = 0; i < count; i++)
            {
                int record = name + 6 + 12 * i;
                if (U16(record) == 3 && U16(record + 2) == 1 && U16(record + 4) == 0x409 && U16(record + 6) <= 6)
                    records.Add(new int[] { U16(record + 6), strings + U16(record + 10), U16(record + 8) });
            }
            if (records.Count == 0) return null;

            int headerSize = 6 + 12 * records.Count;
            MemoryStream text = new MemoryStream();
            byte[] header = new byte[headerSize];
            PutU16(header, 2, records.Count);
            PutU16(header, 4, headerSize);
            for (int i = 0; i < records.Count; i++)
            {
                int at = 6 + 12 * i;
                PutU16(header, at, 3);
                PutU16(header, at + 2, 1);
                PutU16(header, at + 4, 0x409);
                PutU16(header, at + 6, records[i][0]);
                PutU16(header, at + 8, records[i][2]);
                PutU16(header, at + 10, (int)text.Length);
                text.Write(data, records[i][1], records[i][2]);
            }
            byte[] result = new byte[headerSize + text.Length];
            Buffer.BlockCopy(header, 0, result, 0, headerSize);
            Buffer.BlockCopy(text.ToArray(), 0, result, headerSize, (int)text.Length);
            return result;
        }

        static byte[] Assemble(SortedDictionary<string, byte[]> output)
        {
            int count = output.Count;
            int power = 1, log = 0;
            while (power * 2 <= count) { power *= 2; log++; }
            int offset = 12 + 16 * count;
            int total = offset;
            foreach (byte[] table in output.Values) total += (table.Length + 3) & ~3;

            byte[] font = new byte[total];
            PutU32(font, 0, 0x00010000);
            PutU16(font, 4, count);
            PutU16(font, 6, 16 * power);
            PutU16(font, 8, log);
            PutU16(font, 10, 16 * count - 16 * power);
            int index = 0;
            foreach (KeyValuePair<string, byte[]> table in output)
            {
                int record = 12 + 16 * index++;
                for (int i = 0; i < 4; i++) font[record + i] = (byte)table.Key[i];
                Buffer.BlockCopy(table.Value, 0, font, offset, table.Value.Length);
                PutU32(font, record + 4, Checksum(font, offset, (table.Value.Length + 3) & ~3));
                PutU32(font, record + 8, (uint)offset);
                PutU32(font, record + 12, (uint)table.Value.Length);
                offset += (table.Value.Length + 3) & ~3;
            }
            return font;
        }

        static int FindTableOffset(byte[] font, string tag)
        {
            int count = (font[4] << 8) | font[5];
            for (int i = 0; i < count; i++)
            {
                int record = 12 + 16 * i;
                if (Encoding.ASCII.GetString(font, record, 4) == tag)
                    return (int)(((uint)font[record + 8] << 24) | ((uint)font[record + 9] << 16) | ((uint)font[record + 10] << 8) | font[record + 11]);
            }
            throw new InvalidDataException("Font table missing: " + tag);
        }

        static uint Checksum(byte[] bytes, int offset, int length)
        {
            uint sum = 0;
            for (int i = 0; i < length; i += 4)
            {
                uint word = 0;
                for (int j = 0; j < 4; j++)
                    word = (word << 8) | (offset + i + j < bytes.Length && i + j < length ? bytes[offset + i + j] : (byte)0);
                sum = unchecked(sum + word);
            }
            return sum;
        }

        byte[] Copy(string tag)
        {
            int[] entry = tables[tag];
            byte[] result = new byte[entry[1]];
            Buffer.BlockCopy(data, entry[0], result, 0, entry[1]);
            return result;
        }

        int U16(int at) { return (data[at] << 8) | data[at + 1]; }
        int S16(int at) { return (short)U16(at); }
        long U32(int at) { return ((long)data[at] << 24) | ((long)data[at + 1] << 16) | ((long)data[at + 2] << 8) | data[at + 3]; }

        static void PutU16(byte[] bytes, int at, int value)
        {
            bytes[at] = (byte)(value >> 8);
            bytes[at + 1] = (byte)value;
        }

        static void PutU32(byte[] bytes, int at, uint value)
        {
            bytes[at] = (byte)(value >> 24);
            bytes[at + 1] = (byte)(value >> 16);
            bytes[at + 2] = (byte)(value >> 8);
            bytes[at + 3] = (byte)value;
        }
    }
}
