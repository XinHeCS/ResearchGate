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
    public const string EntityTemplatePath = @"Assets\TableImporter\Template\ExcelAssetEntityTemplete.cs.txt";
    public const string ScriptableTemplatePath = @"Assets\TableImporter\Template\ExcelAssetScriptTemplete.cs.txt";

    // Compile constants
    public const string EntityPrefix = "Entity_";
    public const string TablePrefix = "Table_";
}
