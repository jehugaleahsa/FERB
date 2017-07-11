using System;

namespace FERB
{
    [Flags]
    public enum Border
    {
        Top = 1,
        Bottom = 2,
        Left = 4,
        Right = 8,
        Outline = Top | Bottom | Left | Right
    }
}
