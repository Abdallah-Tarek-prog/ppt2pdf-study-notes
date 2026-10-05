// Minimal PDF reader and writer: enough to read the PDF PowerPoint exports,
// change some pages, and write the whole document back out.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Ppt2Pdf
{
    abstract class PdfObject { }

    sealed class PdfNull : PdfObject
    {
        public static readonly PdfNull Instance = new PdfNull();
    }

    sealed class PdfBool : PdfObject
    {
        public readonly bool Value;
        public PdfBool(bool value) { Value = value; }
    }

    sealed class PdfNumber : PdfObject
    {
        public readonly double Value;
        public readonly string Text;

        public PdfNumber(double value) { Value = value; Text = Format(value); }
        public PdfNumber(double value, string text) { Value = value; Text = text; }

        public static string Format(double value)
        {
            double rounded = Math.Round(value, 4);
            if (rounded == 0) return "0";
            return rounded.ToString("0.####", CultureInfo.InvariantCulture);
        }
    }

    sealed class PdfName : PdfObject
    {
        public readonly string Value;
        public PdfName(string value) { Value = value; }
    }

    sealed class PdfString : PdfObject
    {
        // The token exactly as it appears in the file, delimiters included.
        public readonly byte[] Raw;
        public PdfString(byte[] raw) { Raw = raw; }
    }

    sealed class PdfArray : PdfObject
    {
        public readonly List<PdfObject> Items = new List<PdfObject>();
    }

    sealed class PdfDict : PdfObject
    {
        readonly List<string> keys = new List<string>();
        readonly Dictionary<string, PdfObject> values = new Dictionary<string, PdfObject>();

        public IList<string> Keys { get { return keys; } }

        public PdfObject Get(string key)
        {
            PdfObject value;
            return values.TryGetValue(key, out value) ? value : null;
        }

        public void Set(string key, PdfObject value)
        {
            if (!values.ContainsKey(key)) keys.Add(key);
            values[key] = value;
        }

        public void Remove(string key)
        {
            if (values.Remove(key)) keys.Remove(key);
        }

        public PdfDict Clone()
        {
            PdfDict copy = new PdfDict();
            foreach (string key in keys) copy.Set(key, values[key]);
            return copy;
        }
    }

    sealed class PdfRef : PdfObject
    {
        public readonly int Number;
        public readonly int Generation;
        public PdfRef(int number, int generation) { Number = number; Generation = generation; }
    }

    sealed class PdfStream : PdfObject
    {
        public readonly PdfDict Dict;
        public readonly byte[] Data;   // still encoded, as stored in the file
        public PdfStream(PdfDict dict, byte[] data) { Dict = dict; Data = data; }
    }

    // A bare word such as "obj" or "R"; only seen while parsing.
    sealed class PdfKeyword : PdfObject
    {
        public readonly string Value;
        public PdfKeyword(string value) { Value = value; }
    }

    sealed class PdfException : Exception
    {
        public PdfException(string message) : base(message) { }
    }

    sealed class PdfParser
    {
        readonly byte[] data;
        readonly PdfDocument document;
        public int Position;

        public PdfParser(byte[] data, int position, PdfDocument document)
        {
            this.data = data;
            Position = position;
            this.document = document;
        }

        public static bool IsWhitespace(int c)
        {
            return c == 0 || c == 9 || c == 10 || c == 12 || c == 13 || c == 32;
        }

        public static bool IsDelimiter(int c)
        {
            return c == '(' || c == ')' || c == '<' || c == '>' || c == '[' || c == ']' ||
                   c == '{' || c == '}' || c == '/' || c == '%';
        }

        public void SkipWhitespace()
        {
            while (Position < data.Length)
            {
                int c = data[Position];
                if (IsWhitespace(c))
                {
                    Position++;
                }
                else if (c == '%')
                {
                    while (Position < data.Length && data[Position] != '\n' && data[Position] != '\r') Position++;
                }
                else
                {
                    break;
                }
            }
        }

        public PdfObject ReadObject()
        {
            SkipWhitespace();
            if (Position >= data.Length) throw new PdfException("Unexpected end of PDF data.");
            int c = data[Position];
            if (c == '/') return ReadName();
            if (c == '(') return ReadLiteralString();
            if (c == '<')
            {
                if (Position + 1 < data.Length && data[Position + 1] == '<') return ReadDictionary();
                return ReadHexString();
            }
            if (c == '[') return ReadArray();
            if (c == '+' || c == '-' || c == '.' || (c >= '0' && c <= '9')) return ReadNumberOrReference();

            string word = ReadWord();
            if (word.Length == 0) throw new PdfException("Unexpected character in PDF at offset " + Position + ".");
            if (word == "true") return new PdfBool(true);
            if (word == "false") return new PdfBool(false);
            if (word == "null") return PdfNull.Instance;
            return new PdfKeyword(word);
        }

        string ReadWord()
        {
            int start = Position;
            while (Position < data.Length && !IsWhitespace(data[Position]) && !IsDelimiter(data[Position])) Position++;
            return Encoding.ASCII.GetString(data, start, Position - start);
        }

        PdfName ReadName()
        {
            Position++;
            StringBuilder name = new StringBuilder();
            while (Position < data.Length && !IsWhitespace(data[Position]) && !IsDelimiter(data[Position]))
            {
                int c = data[Position];
                if (c == '#' && Position + 2 < data.Length && IsHex(data[Position + 1]) && IsHex(data[Position + 2]))
                {
                    name.Append((char)(HexValue(data[Position + 1]) * 16 + HexValue(data[Position + 2])));
                    Position += 3;
                }
                else
                {
                    name.Append((char)c);
                    Position++;
                }
            }
            return new PdfName(name.ToString());
        }

        PdfString ReadLiteralString()
        {
            int start = Position;
            Position++;
            int depth = 1;
            while (Position < data.Length && depth > 0)
            {
                int c = data[Position++];
                if (c == '\\') Position++;
                else if (c == '(') depth++;
                else if (c == ')') depth--;
            }
            return new PdfString(Slice(start, Math.Min(Position, data.Length)));
        }

        PdfString ReadHexString()
        {
            int start = Position;
            while (Position < data.Length && data[Position] != '>') Position++;
            Position = Math.Min(Position + 1, data.Length);
            return new PdfString(Slice(start, Position));
        }

        PdfDict ReadDictionary()
        {
            Position += 2;
            PdfDict dict = new PdfDict();
            while (true)
            {
                SkipWhitespace();
                if (Position >= data.Length) throw new PdfException("Unterminated dictionary in PDF.");
                if (data[Position] == '>' && Position + 1 < data.Length && data[Position + 1] == '>')
                {
                    Position += 2;
                    return dict;
                }
                PdfName key = ReadObject() as PdfName;
                if (key == null) throw new PdfException("Dictionary key is not a name at offset " + Position + ".");
                dict.Set(key.Value, ReadObject());
            }
        }

        PdfArray ReadArray()
        {
            Position++;
            PdfArray array = new PdfArray();
            while (true)
            {
                SkipWhitespace();
                if (Position >= data.Length) throw new PdfException("Unterminated array in PDF.");
                if (data[Position] == ']')
                {
                    Position++;
                    return array;
                }
                array.Items.Add(ReadObject());
            }
        }

        PdfObject ReadNumberOrReference()
        {
            string text = ReadWord();
            double value;
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) value = 0;
            if (text.IndexOf('.') >= 0 || text[0] == '+' || text[0] == '-') return new PdfNumber(value, text);

            int saved = Position;
            SkipWhitespace();
            if (Position < data.Length && data[Position] >= '0' && data[Position] <= '9')
            {
                string generation = ReadWord();
                SkipWhitespace();
                int g;
                if (Position < data.Length && data[Position] == 'R' &&
                    (Position + 1 >= data.Length || IsWhitespace(data[Position + 1]) || IsDelimiter(data[Position + 1])) &&
                    int.TryParse(generation, NumberStyles.None, CultureInfo.InvariantCulture, out g))
                {
                    Position++;
                    return new PdfRef((int)value, g);
                }
            }
            Position = saved;
            return new PdfNumber(value, text);
        }

        public long ReadInteger()
        {
            SkipWhitespace();
            string text = ReadWord();
            long value;
            if (!long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value))
                throw new PdfException("Expected a number in PDF at offset " + Position + ".");
            return value;
        }

        public bool TryKeyword(string keyword)
        {
            SkipWhitespace();
            int end = Position + keyword.Length;
            if (end > data.Length) return false;
            for (int i = 0; i < keyword.Length; i++)
                if (data[Position + i] != keyword[i]) return false;
            if (end < data.Length && !IsWhitespace(data[end]) && !IsDelimiter(data[end])) return false;
            Position = end;
            return true;
        }

        public PdfObject ReadIndirectObject(out int number)
        {
            number = (int)ReadInteger();
            ReadInteger();
            if (!TryKeyword("obj")) throw new PdfException("Missing 'obj' keyword in PDF.");
            PdfObject value = ReadObject();
            PdfDict dict = value as PdfDict;
            if (dict == null || !TryKeyword("stream")) return value;

            if (Position < data.Length && data[Position] == '\r') Position++;
            if (Position < data.Length && data[Position] == '\n') Position++;
            int start = Position;
            int length = -1;
            PdfObject lengthObject = document != null ? document.Resolve(dict.Get("Length")) : dict.Get("Length");
            if (lengthObject is PdfNumber) length = (int)((PdfNumber)lengthObject).Value;
            if (length < 0 || start + length > data.Length || !EndstreamFollows(start + length))
                length = FindEndstream(start) - start;
            Position = start + length;
            TryKeyword("endstream");
            return new PdfStream(dict, Slice(start, start + length));
        }

        bool EndstreamFollows(int offset)
        {
            int saved = Position;
            Position = offset;
            bool found = TryKeyword("endstream");
            Position = saved;
            return found;
        }

        int FindEndstream(int start)
        {
            byte[] marker = Encoding.ASCII.GetBytes("endstream");
            for (int i = start; i + marker.Length <= data.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < marker.Length && match; j++) match = data[i + j] == marker[j];
                if (!match) continue;
                int end = i;
                if (end > start && data[end - 1] == '\n') end--;
                if (end > start && data[end - 1] == '\r') end--;
                return end;
            }
            throw new PdfException("Unterminated stream in PDF.");
        }

        byte[] Slice(int start, int end)
        {
            byte[] result = new byte[end - start];
            Buffer.BlockCopy(data, start, result, 0, result.Length);
            return result;
        }

        static bool IsHex(int c)
        {
            return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
        }

        static int HexValue(int c)
        {
            if (c <= '9') return c - '0';
            if (c <= 'F') return c - 'A' + 10;
            return c - 'a' + 10;
        }
    }

    sealed class PdfDocument
    {
        sealed class XrefEntry
        {
            public long Offset;          // for objects stored directly in the file
            public int StreamNumber = -1; // for objects stored inside an object stream
            public int Index;
        }

        sealed class ObjectStream
        {
            public byte[] Data;
            public int[] Numbers;
            public int[] Offsets;
        }

        public readonly byte[] Data;
        public PdfDict Trailer;
        readonly Dictionary<int, XrefEntry> xref = new Dictionary<int, XrefEntry>();
        readonly Dictionary<int, PdfObject> cache = new Dictionary<int, PdfObject>();
        readonly Dictionary<int, ObjectStream> objectStreams = new Dictionary<int, ObjectStream>();
        readonly HashSet<int> loading = new HashSet<int>();

        PdfDocument(byte[] data) { Data = data; }

        public static PdfDocument Load(byte[] data)
        {
            PdfDocument document = new PdfDocument(data);
            document.ReadCrossReferences();
            if (document.Trailer == null || document.Trailer.Get("Root") == null)
                throw new PdfException("The PDF has no document catalog.");
            return document;
        }

        public int MaxObjectNumber
        {
            get
            {
                int max = 0;
                foreach (int number in xref.Keys) max = Math.Max(max, number);
                PdfNumber size = Trailer.Get("Size") as PdfNumber;
                if (size != null) max = Math.Max(max, (int)size.Value - 1);
                return max;
            }
        }

        public string Version
        {
            get
            {
                string head = Encoding.ASCII.GetString(Data, 0, Math.Min(16, Data.Length));
                int at = head.IndexOf("%PDF-", StringComparison.Ordinal);
                if (at < 0 || at + 8 > head.Length) return "1.7";
                return head.Substring(at + 5, 3);
            }
        }

        public PdfObject Resolve(PdfObject value)
        {
            int guard = 0;
            while (value is PdfRef && guard++ < 32) value = GetObject(((PdfRef)value).Number);
            return value is PdfRef ? null : value;
        }

        public PdfObject GetObject(int number)
        {
            PdfObject value;
            if (cache.TryGetValue(number, out value)) return value;
            XrefEntry entry;
            if (!xref.TryGetValue(number, out entry) || !loading.Add(number)) return null;
            try
            {
                if (entry.StreamNumber >= 0)
                {
                    value = ReadFromObjectStream(number, entry);
                }
                else
                {
                    int found;
                    value = new PdfParser(Data, (int)entry.Offset, this).ReadIndirectObject(out found);
                    if (found != number) throw new PdfException("PDF cross-reference table points to the wrong object.");
                }
            }
            finally
            {
                loading.Remove(number);
            }
            cache[number] = value;
            return value;
        }

        public PdfObject GetInherited(PdfDict page, string key)
        {
            PdfDict node = page;
            for (int depth = 0; node != null && depth < 64; depth++)
            {
                PdfObject value = node.Get(key);
                if (value != null) return value;
                node = Resolve(node.Get("Parent")) as PdfDict;
            }
            return null;
        }

        public List<PdfRef> GetPages()
        {
            List<PdfRef> pages = new List<PdfRef>();
            PdfDict catalog = Resolve(Trailer.Get("Root")) as PdfDict;
            if (catalog == null) throw new PdfException("The PDF has no document catalog.");
            CollectPages(catalog.Get("Pages"), pages, new HashSet<int>());
            return pages;
        }

        void CollectPages(PdfObject node, List<PdfRef> pages, HashSet<int> visited)
        {
            PdfRef reference = node as PdfRef;
            if (reference == null || !visited.Add(reference.Number)) return;
            PdfDict dict = Resolve(reference) as PdfDict;
            if (dict == null) return;
            PdfName type = dict.Get("Type") as PdfName;
            PdfArray kids = Resolve(dict.Get("Kids")) as PdfArray;
            if (kids != null && (type == null || type.Value == "Pages"))
            {
                foreach (PdfObject kid in kids.Items) CollectPages(kid, pages, visited);
            }
            else
            {
                pages.Add(reference);
            }
        }

        void ReadCrossReferences()
        {
            long offset = FindStartXref();
            HashSet<long> visited = new HashSet<long>();
            while (offset >= 0 && offset < Data.Length && visited.Add(offset))
            {
                PdfDict trailer = ReadXrefSection(offset);
                if (Trailer == null) Trailer = trailer;
                PdfNumber previous = trailer.Get("Prev") as PdfNumber;
                offset = previous != null ? (long)previous.Value : -1;
            }
        }

        long FindStartXref()
        {
            byte[] marker = Encoding.ASCII.GetBytes("startxref");
            int lowest = Math.Max(0, Data.Length - 4096);
            for (int i = Data.Length - marker.Length; i >= lowest; i--)
            {
                bool match = true;
                for (int j = 0; j < marker.Length && match; j++) match = Data[i + j] == marker[j];
                if (match) return new PdfParser(Data, i + marker.Length, null).ReadInteger();
            }
            throw new PdfException("The PDF has no cross-reference table.");
        }

        // Newer sections are read first, so entries already present are kept.
        PdfDict ReadXrefSection(long offset)
        {
            PdfParser parser = new PdfParser(Data, (int)offset, null);
            if (!parser.TryKeyword("xref"))
            {
                return ReadXrefStream(offset);
            }

            List<KeyValuePair<int, long>> entries = new List<KeyValuePair<int, long>>();
            while (!parser.TryKeyword("trailer"))
            {
                long first = parser.ReadInteger();
                long count = parser.ReadInteger();
                for (long i = 0; i < count; i++)
                {
                    long entryOffset = parser.ReadInteger();
                    parser.ReadInteger();
                    parser.SkipWhitespace();
                    bool inUse = parser.TryKeyword("n");
                    if (!inUse && !parser.TryKeyword("f")) throw new PdfException("Malformed PDF cross-reference table.");
                    if (inUse && entryOffset > 0) entries.Add(new KeyValuePair<int, long>((int)(first + i), entryOffset));
                }
            }
            PdfDict trailer = parser.ReadObject() as PdfDict;
            if (trailer == null) throw new PdfException("Malformed PDF trailer.");

            // Hybrid files keep some objects in a cross-reference stream as well.
            PdfNumber hidden = trailer.Get("XRefStm") as PdfNumber;
            if (hidden != null) ReadXrefStream((long)hidden.Value);
            foreach (KeyValuePair<int, long> entry in entries)
                if (!xref.ContainsKey(entry.Key)) xref[entry.Key] = new XrefEntry { Offset = entry.Value };
            return trailer;
        }

        PdfDict ReadXrefStream(long offset)
        {
            int number;
            PdfStream stream = new PdfParser(Data, (int)offset, null).ReadIndirectObject(out number) as PdfStream;
            if (stream == null) throw new PdfException("Malformed PDF cross-reference stream.");
            byte[] data = Decode(stream);
            PdfArray widthsArray = stream.Dict.Get("W") as PdfArray;
            if (widthsArray == null || widthsArray.Items.Count < 3) throw new PdfException("Malformed PDF cross-reference stream.");
            int[] widths = new int[3];
            for (int i = 0; i < 3; i++) widths[i] = (int)((PdfNumber)widthsArray.Items[i]).Value;

            List<int> ranges = new List<int>();
            PdfArray index = stream.Dict.Get("Index") as PdfArray;
            if (index != null)
            {
                foreach (PdfObject item in index.Items) ranges.Add((int)((PdfNumber)item).Value);
            }
            else
            {
                ranges.Add(0);
                ranges.Add((int)((PdfNumber)stream.Dict.Get("Size")).Value);
            }

            int position = 0;
            int rowLength = widths[0] + widths[1] + widths[2];
            for (int r = 0; r + 1 < ranges.Count; r += 2)
            {
                for (int i = 0; i < ranges[r + 1] && position + rowLength <= data.Length; i++)
                {
                    long type = widths[0] == 0 ? 1 : ReadField(data, ref position, widths[0]);
                    long field2 = ReadField(data, ref position, widths[1]);
                    long field3 = ReadField(data, ref position, widths[2]);
                    int objectNumber = ranges[r] + i;
                    if (xref.ContainsKey(objectNumber)) continue;
                    if (type == 1 && field2 > 0) xref[objectNumber] = new XrefEntry { Offset = field2 };
                    else if (type == 2) xref[objectNumber] = new XrefEntry { StreamNumber = (int)field2, Index = (int)field3 };
                }
            }
            return stream.Dict;
        }

        static long ReadField(byte[] data, ref int position, int width)
        {
            long value = 0;
            for (int i = 0; i < width; i++) value = (value << 8) | data[position++];
            return value;
        }

        PdfObject ReadFromObjectStream(int number, XrefEntry entry)
        {
            ObjectStream objectStream;
            if (!objectStreams.TryGetValue(entry.StreamNumber, out objectStream))
            {
                PdfStream stream = GetObject(entry.StreamNumber) as PdfStream;
                if (stream == null) throw new PdfException("Missing PDF object stream.");
                objectStream = new ObjectStream();
                objectStream.Data = Decode(stream);
                int count = (int)((PdfNumber)stream.Dict.Get("N")).Value;
                int first = (int)((PdfNumber)stream.Dict.Get("First")).Value;
                objectStream.Numbers = new int[count];
                objectStream.Offsets = new int[count];
                PdfParser header = new PdfParser(objectStream.Data, 0, null);
                for (int i = 0; i < count; i++)
                {
                    objectStream.Numbers[i] = (int)header.ReadInteger();
                    objectStream.Offsets[i] = first + (int)header.ReadInteger();
                }
                objectStreams[entry.StreamNumber] = objectStream;
            }

            int index = entry.Index;
            if (index < 0 || index >= objectStream.Numbers.Length || objectStream.Numbers[index] != number)
                index = Array.IndexOf(objectStream.Numbers, number);
            if (index < 0) return null;
            return new PdfParser(objectStream.Data, objectStream.Offsets[index], this).ReadObject();
        }

        public byte[] Decode(PdfStream stream)
        {
            PdfObject filter = Resolve(stream.Dict.Get("Filter"));
            PdfObject parameters = Resolve(stream.Dict.Get("DecodeParms"));
            List<string> filters = new List<string>();
            List<PdfObject> parameterList = new List<PdfObject>();
            if (filter is PdfName)
            {
                filters.Add(((PdfName)filter).Value);
                parameterList.Add(parameters);
            }
            else if (filter is PdfArray)
            {
                PdfArray filterArray = (PdfArray)filter;
                PdfArray parameterArray = parameters as PdfArray;
                for (int i = 0; i < filterArray.Items.Count; i++)
                {
                    filters.Add(((PdfName)Resolve(filterArray.Items[i])).Value);
                    parameterList.Add(parameterArray != null && i < parameterArray.Items.Count ? Resolve(parameterArray.Items[i]) : null);
                }
            }

            byte[] data = stream.Data;
            for (int i = 0; i < filters.Count; i++)
            {
                if (filters[i] != "FlateDecode" && filters[i] != "Fl")
                    throw new PdfException("Unsupported PDF filter: " + filters[i] + ".");
                data = Unpredict(Zlib.Inflate(data), parameterList[i] as PdfDict);
            }
            return data;
        }

        static byte[] Unpredict(byte[] data, PdfDict parameters)
        {
            if (parameters == null) return data;
            PdfNumber predictorNumber = parameters.Get("Predictor") as PdfNumber;
            int predictor = predictorNumber != null ? (int)predictorNumber.Value : 1;
            if (predictor < 10) return data;   // 1 means none; TIFF (2) is not used by PowerPoint

            int columns = NumberOr(parameters, "Columns", 1);
            int colors = NumberOr(parameters, "Colors", 1);
            int bits = NumberOr(parameters, "BitsPerComponent", 8);
            int pixelBytes = Math.Max(1, colors * bits / 8);
            int rowBytes = (colors * bits * columns + 7) / 8;
            MemoryStream output = new MemoryStream();
            byte[] previous = new byte[rowBytes];
            byte[] row = new byte[rowBytes];
            for (int position = 0; position < data.Length; position += rowBytes + 1)
            {
                int type = data[position];
                int available = Math.Min(rowBytes, data.Length - position - 1);
                Array.Clear(row, 0, rowBytes);
                Buffer.BlockCopy(data, position + 1, row, 0, available);
                for (int i = 0; i < rowBytes; i++)
                {
                    int left = i >= pixelBytes ? row[i - pixelBytes] : 0;
                    int up = previous[i];
                    int upLeft = i >= pixelBytes ? previous[i - pixelBytes] : 0;
                    switch (type)
                    {
                        case 1: row[i] = (byte)(row[i] + left); break;
                        case 2: row[i] = (byte)(row[i] + up); break;
                        case 3: row[i] = (byte)(row[i] + (left + up) / 2); break;
                        case 4: row[i] = (byte)(row[i] + Paeth(left, up, upLeft)); break;
                    }
                }
                output.Write(row, 0, available);
                byte[] swap = previous;
                previous = row;
                row = swap;
            }
            return output.ToArray();
        }

        static int Paeth(int a, int b, int c)
        {
            int p = a + b - c;
            int pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
            if (pa <= pb && pa <= pc) return a;
            return pb <= pc ? b : c;
        }

        static int NumberOr(PdfDict dict, string key, int fallback)
        {
            PdfNumber number = dict.Get(key) as PdfNumber;
            return number != null ? (int)number.Value : fallback;
        }
    }

    static class Zlib
    {
        public static byte[] Inflate(byte[] data)
        {
            int start = data.Length >= 2 && (data[0] & 0x0F) == 8 && ((data[0] << 8) | data[1]) % 31 == 0 ? 2 : 0;
            using (MemoryStream input = new MemoryStream(data, start, data.Length - start))
            using (DeflateStream inflater = new DeflateStream(input, CompressionMode.Decompress))
            using (MemoryStream output = new MemoryStream())
            {
                inflater.CopyTo(output);
                return output.ToArray();
            }
        }

        public static byte[] Deflate(byte[] data)
        {
            using (MemoryStream output = new MemoryStream())
            {
                output.WriteByte(0x78);
                output.WriteByte(0x9C);
                using (DeflateStream deflater = new DeflateStream(output, CompressionLevel.Optimal, true))
                {
                    deflater.Write(data, 0, data.Length);
                }
                uint a = 1, b = 0;
                foreach (byte value in data)
                {
                    a = (a + value) % 65521;
                    b = (b + a) % 65521;
                }
                uint adler = (b << 16) | a;
                output.WriteByte((byte)(adler >> 24));
                output.WriteByte((byte)(adler >> 16));
                output.WriteByte((byte)(adler >> 8));
                output.WriteByte((byte)adler);
                return output.ToArray();
            }
        }
    }

    // Writes a fresh copy of a document: only objects reachable from the
    // catalog are kept, renumbered from 1, with a classic cross-reference table.
    sealed class PdfRewriter
    {
        readonly PdfDocument document;
        readonly Dictionary<int, PdfObject> replacements = new Dictionary<int, PdfObject>();
        readonly Dictionary<int, int> newNumbers = new Dictionary<int, int>();
        readonly List<int> order = new List<int>();
        int nextNumber;
        MemoryStream output;

        public PdfRewriter(PdfDocument document)
        {
            this.document = document;
            nextNumber = document.MaxObjectNumber + 1;
        }

        public PdfRef Add(PdfObject value)
        {
            PdfRef reference = Reserve();
            replacements[reference.Number] = value;
            return reference;
        }

        public PdfRef Reserve()
        {
            return new PdfRef(nextNumber++, 0);
        }

        public void Replace(int number, PdfObject value)
        {
            replacements[number] = value;
        }

        public byte[] Write(PdfDict trailer)
        {
            output = new MemoryStream();
            Ascii("%PDF-" + document.Version + "\n%âãÏÓ\n");

            PdfDict newTrailer = new PdfDict();
            foreach (string key in trailer.Keys)
            {
                if (key == "Prev" || key == "XRefStm" || key == "Size" || key == "Encrypt") continue;
                if (key != "Root" && key != "Info" && key != "ID") continue;
                newTrailer.Set(key, trailer.Get(key));
            }
            MemoryStream trailerBytes = new MemoryStream();
            MemoryStream saved = output;
            output = trailerBytes;
            WriteValue(newTrailer);
            output = saved;

            List<long> offsets = new List<long>();
            for (int i = 0; i < order.Count; i++)
            {
                int original = order[i];
                PdfObject value;
                if (!replacements.TryGetValue(original, out value)) value = document.GetObject(original);
                offsets.Add(output.Position);
                Ascii((i + 1).ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
                PdfStream stream = value as PdfStream;
                if (stream != null)
                {
                    PdfDict dict = stream.Dict.Clone();
                    dict.Set("Length", new PdfNumber(stream.Data.Length));
                    WriteValue(dict);
                    Ascii("\nstream\n");
                    output.Write(stream.Data, 0, stream.Data.Length);
                    Ascii("\nendstream");
                }
                else
                {
                    WriteValue(value);
                }
                Ascii("\nendobj\n");
            }

            long xrefOffset = output.Position;
            Ascii("xref\n0 " + (order.Count + 1) + "\n0000000000 65535 f\r\n");
            foreach (long offset in offsets) Ascii(offset.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n\r\n");
            Ascii("trailer\n<</Size " + (order.Count + 1));
            byte[] rest = trailerBytes.ToArray();
            output.Write(rest, 2, rest.Length - 2);   // the remaining entries, without the opening "<<"
            Ascii("\nstartxref\n" + xrefOffset + "\n%%EOF\n");
            return output.ToArray();
        }

        int Map(int original)
        {
            int number;
            if (!newNumbers.TryGetValue(original, out number))
            {
                order.Add(original);
                number = order.Count;
                newNumbers[original] = number;
            }
            return number;
        }

        void WriteValue(PdfObject value)
        {
            if (value == null || value is PdfNull)
            {
                Ascii("null");
            }
            else if (value is PdfBool)
            {
                Ascii(((PdfBool)value).Value ? "true" : "false");
            }
            else if (value is PdfNumber)
            {
                Ascii(((PdfNumber)value).Text);
            }
            else if (value is PdfName)
            {
                WriteName(((PdfName)value).Value);
            }
            else if (value is PdfString)
            {
                byte[] raw = ((PdfString)value).Raw;
                output.Write(raw, 0, raw.Length);
            }
            else if (value is PdfRef)
            {
                Ascii(Map(((PdfRef)value).Number) + " 0 R");
            }
            else if (value is PdfArray)
            {
                Ascii("[");
                bool first = true;
                foreach (PdfObject item in ((PdfArray)value).Items)
                {
                    if (!first) Ascii(" ");
                    WriteValue(item);
                    first = false;
                }
                Ascii("]");
            }
            else if (value is PdfDict)
            {
                PdfDict dict = (PdfDict)value;
                Ascii("<<");
                foreach (string key in dict.Keys)
                {
                    WriteName(key);
                    Ascii(" ");
                    WriteValue(dict.Get(key));
                }
                Ascii(">>");
            }
            else if (value is PdfKeyword)
            {
                Ascii(((PdfKeyword)value).Value);
            }
            else
            {
                throw new PdfException("A stream cannot be stored inside another PDF object.");
            }
        }

        void WriteName(string name)
        {
            StringBuilder text = new StringBuilder("/");
            foreach (char c in name)
            {
                if (c < 0x21 || c > 0x7E || c == '#' || PdfParser.IsDelimiter(c))
                    text.Append('#').Append(((int)c & 0xFF).ToString("X2"));
                else
                    text.Append(c);
            }
            Ascii(text.ToString());
        }

        void Ascii(string text)
        {
            byte[] bytes = new byte[text.Length];
            for (int i = 0; i < text.Length; i++) bytes[i] = (byte)text[i];
            output.Write(bytes, 0, bytes.Length);
        }
    }
}
