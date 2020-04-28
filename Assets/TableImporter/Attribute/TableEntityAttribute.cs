using System;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
public class TableEntityAttribute : Attribute
{
    public string[] FieldNames { get; set; }
    public string[] FieldTypes { get; set; }
}
