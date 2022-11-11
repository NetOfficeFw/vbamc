// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;
using vbamc.Vba;

namespace vbamc
{
    public class Guard
    {
        public static void EnsureLengthBetween(byte[] value, int min, int max, string paramName, string message)
        {
            if (!(value.Length >= min && value.Length <= max))
            {
                throw new ArgumentOutOfRangeException(paramName, message);
            }
        }

        public static void EnsureNoNullCharacters(byte[] value, string paramName, string message)
        {
            if (value.Any(b => b == VbaEncodings.NULL))
            {
                throw new ArgumentOutOfRangeException(paramName, message);
            }
        }        
    }
}
