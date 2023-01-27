// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using System.Text;

namespace vbamc.Vba
{
    public static class VbaEncodings
    {
        public const int DefaultCodePage = 1252;
        public const byte NULL = 0x00;
        public static readonly byte[] NewLine = new byte[] { 0x0D, 0x0A };

        public static readonly Encoding Default = Encoding.GetEncoding(DefaultCodePage);
        public static readonly Encoding UTF16 = Encoding.Unicode;

        public static void WriteLine(this BinaryWriter writer, string value)
        {
            var bytes = Default.GetBytes(value);
            writer.Write(bytes);
            writer.Write(NewLine);
        }

        public static void WriteLine(this BinaryWriter writer)
        {
            writer.Write(NewLine);
        }
    }
}
