using System;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class TableAssetAttribute : Attribute
{
    public string AssetPath { get; set; }
    public string ExcelName { get; set; }
    public bool LogOnImport { get; set; }
}
