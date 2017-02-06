// Guids.cs
// MUST match guids.h
using System;
// ReSharper disable InconsistentNaming

namespace BlackIceSoftware.CBExtensionPkg
{
    static class GuidList
    {
        public const string guidCBExtensionPkgPkgString = "b467e5f1-4a69-4ad8-a2a6-0d8bf3932e0e";
        public const string guidCBExtensionPkgCmdSetString = "3001e580-502a-454e-9ff4-9f70884be3d8";

        public static readonly Guid guidCBExtensionPkgCmdSet = new Guid(guidCBExtensionPkgCmdSetString);
    };
}