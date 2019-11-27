using System;

namespace MMCSirUtilities
{
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Parameter, AllowMultiple = false)]
    public class ExcludePropertyFromDataTable : Attribute { }
}