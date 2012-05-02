// Guids.cs
// MUST match guids.h
using System;

namespace ScottHanselman.ReloadPackage
{
    static class GuidList
    {
        public const string guidReloadPackagePkgString = "e435087d-33d4-4f94-b465-f4c163476b71";
        public const string guidReloadPackageCmdSetString = "bbc95ddc-953b-4c54-952e-e06720e31b95";

        public static readonly Guid guidReloadPackageCmdSet = new Guid(guidReloadPackageCmdSetString);
    };
}