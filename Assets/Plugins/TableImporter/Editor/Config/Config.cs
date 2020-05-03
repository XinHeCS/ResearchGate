using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class Config
{
    // path to store the table scriptableobjects
    public const string TableScriptablePath = @"Assets\ConfigTableData\TableScriptable";
    public const string TableEntityPath = @"Assets\ConfigTableData\TableEntity";
    public const string TableDataPath = @"Assets\ConfigTableData\TableData";

    // Template file path
    //public const string EntityTemplatePath = @"Assets\Plugins\TableImporter\Template\ExcelAssetEntityTemplete.cs.txt";
    public const string ScriptableTemplatePath = @"Assets\Plugins\TableImporter\Template\ExcelAssetScriptTemplete.cs.txt";

    // Compile constants
    public const string EntityPrefix = "Entity_";
    public const string TablePrefix = "Table_";
    public const string ScriptablePrefix = "Scriptable_";

    // Reference config table objects assembly name
    public const string DefaultAssembly = "Assembly-CSharp";
    public const string TableAssemblyName = "ConfigTableAssembly";

    // Flag to trigger log on import
    public const bool LogOnImport = true;
}
