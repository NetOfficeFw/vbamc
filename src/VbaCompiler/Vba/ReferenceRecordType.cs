// Copyright 2022 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

using System;

namespace vbamc.Vba
{
    public enum ReferenceRecordType : ushort
    {
        Control = 0x002F,
        Original = 0x0033,
        Registered = 0x000D,
        Project = 0x000E
    }
}
