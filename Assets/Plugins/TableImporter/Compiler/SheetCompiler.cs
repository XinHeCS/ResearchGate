using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

public class SheetCompiler
{
    class TypeDescription
    {
        public string _typeName; // Type name
        public string _fieldName;
        public string[] _values; // Value domain for special type (e.g. enum)
    }

    private List<TypeDescription> _typeDescriptions; // For definition of enum field in sheet
    private TableCompiler _tableCompiler;
    private ISheet _sheet;

    // Flag to control whether stop the compiling process
    private bool _isStop = false;
    // Flag to indicate whether the compile process success
    private bool _isSucess = true;

    // Error messages of compile process
    private string _errMsg;

    public string ErrorMessage
    {
        get
        {
            return _errMsg;
        }
    }

    public bool CompileSuccess
    {
        get
        {
            return _isSucess;
        }
    }

    public SheetCompiler(ISheet sheet, TableCompiler compiler)
    {
        _sheet = sheet;
        _tableCompiler = compiler;
    }

    public void Compile()
    {
        ParseSheet();
        if (_isStop)
        {
            return;
        }
        GenerateScript();
    }

    public void ParseSheet()
    {
        ParseFieldName();
        ParseFieldType();
    }

    private void ParseFieldName()
    {
        List<string> fieldNames = new List<string>();
        IRow headerRow = _sheet.GetRow(0);
        if (headerRow == null)
        {
            _isStop = true;
            _isSucess = true;
            return;
        }
        for (int i = 0; i <= headerRow.LastCellNum; ++i)
        {
            ICell cell = headerRow.GetCell(i);
            if (cell == null || cell.CellType == CellType.Blank)
            {
                break;
            }
            else if (cell.CellType == CellType.String)
            {
                fieldNames.Add(cell.StringCellValue.Trim());
            }
            else if (cell.CellType == CellType.Unknown)
            {
                StopCompile(string.Format("Error in compiling sheet {0} ( {1} : {2} ): Unknown header is not allowed",
                    _sheet.SheetName, cell.RowIndex, i));
                return;
            }
        }
        if (_tableCompiler.FieldNames == null)
        {
            _tableCompiler.FieldNames = fieldNames;
        }
        else if (!CheckHeader(fieldNames))
        {
            StopCompile(string.Format("Error in compiling sheet {0}: Unmatched header fields", _sheet.SheetName));
        }
    }

    private void ParseFieldType()
    {
        List<string> fieldTypes = new List<string>();
        IRow typeRow = _sheet.GetRow(1);
        if (typeRow == null)
        {
            StopCompile(
                string.Format("Error in compiling sheet {0}: Missing type definitoins.", _sheet.SheetName));
            return;
        }
        for (int i = 0; i <= typeRow.LastCellNum; ++i)
        {
            ICell cell = typeRow.GetCell(i);
            if (cell == null || cell.CellType == CellType.Blank)
            {
                break;
            }
            else if (cell.CellType == CellType.String)
            {
                var rawTypeInfo = cell.StringCellValue.Replace(" ", "").ToLower();
                var defInfo = rawTypeInfo.Split('|');
                var type = defInfo[0];
                if (!TableDataType.HasType(type))
                {
                    StopCompile(string.Format("Error in compiling sheet {0} ( {1} : {2} ): Unsupported type {3}",
                        _sheet.SheetName, cell.RowIndex, i, type));
                    return;
                }
                if (type == TableDataType.k_EnumType) // Enum type name needs to handle particulaly
                {
                    type += "_" + _tableCompiler.FieldNames[i];
                }
                ParseTypeInfo(type, _tableCompiler.FieldNames[i], defInfo);
                fieldTypes.Add(rawTypeInfo);
            }
            else if (cell.CellType == CellType.Unknown)
            {
                StopCompile(string.Format("Error in compiling sheet {0} ( {1} : {2} ): Unknown field is not allowed",
                    _sheet.SheetName, cell.RowIndex, i));
                return;
            }
        }
        if (_tableCompiler.FieldTypes == null)
        {
            _tableCompiler.FieldTypes = fieldTypes;
        }
        else if (!CheckTypes(fieldTypes))
        {
            StopCompile(string.Format("Error in compiling sheet {0}: Unmatched field type", _sheet.SheetName));
        }
    }

    private void GenerateScript()
    {     
        StringBuilder sb = new StringBuilder();
        sb.Append(File.ReadAllText(Config.ScriptableTemplatePath));

        // Fill Entity content in template
        // Fill Field names in attribute
        sb.Replace(TemplateSymbol.k_FIELDNAMES, GetFieldNameInAttr().ToString());
        // Fill Field types in attribute
        sb.Replace(TemplateSymbol.k_FIELDTYPES, GetFieldTypeInAttr().ToString());
        // Fill Entity names
        sb.Replace(TemplateSymbol.k_ENTITYNAME, _tableCompiler.EntityName);
        // Fill Entity Body
        sb.Replace(TemplateSymbol.k_FIELDS, GetEntityBody().ToString());
        // Fill extra definitions
        sb.Replace(TemplateSymbol.k_ENUMDEF, GetEnumScript().ToString());

        // Fill table content in template
        var excelPath = _tableCompiler.FilePath;
        var assetName = _tableCompiler.EntityName.Replace(Config.EntityPrefix, Config.ScriptablePrefix);
        sb.Replace(TemplateSymbol.k_EXCELPATH, excelPath);
        sb.Replace(TemplateSymbol.k_ASSETSCRIPTNAME, assetName);

        // Write to disk
        if (!Directory.Exists(Config.TableScriptablePath))
        {
            Directory.CreateDirectory(Config.TableScriptablePath);
        }
        var path = Path.Combine(
            Config.TableScriptablePath,
            assetName + ".cs"
            );
        File.WriteAllText(path, sb.ToString());
    }

    private StringBuilder GetFieldNameInAttr()
    {
        StringBuilder fieldNames = new StringBuilder("\"");
        fieldNames.Append(_tableCompiler.FieldNames[0]);
        fieldNames.Append("\"");
        for (int i = 1; i < _tableCompiler.FieldNames.Count; ++i)
        {
            fieldNames.Append(", \"");
            fieldNames.Append(_tableCompiler.FieldNames[i]);
            fieldNames.Append("\"");
        }

        return fieldNames;
    }

    private StringBuilder GetFieldTypeInAttr()
    {
        StringBuilder fieldTypes = new StringBuilder("\"");
        fieldTypes.Append(_tableCompiler.FieldTypes[0]);
        fieldTypes.Append("\"");
        for (int i = 1; i < _tableCompiler.FieldTypes.Count; ++i)
        {
            fieldTypes.Append(", \"");
            fieldTypes.Append(_tableCompiler.FieldTypes[i]);
            fieldTypes.Append("\"");
        }

        return fieldTypes;
    }

    private StringBuilder GetEntityBody()
    {
        StringBuilder body = new StringBuilder();

        for (int i = 0; i < _typeDescriptions.Count; ++i)
        {
            if (i == 0)
            {
                body.Append("\tpublic ");
            }
            else
            {
                body.Append("\n\tpublic ");
            }
            body.Append(_typeDescriptions[i]._typeName);
            body.Append(" ");
            body.Append(_typeDescriptions[i]._fieldName);
            body.Append(";");
        }

        return body;
    }

    private StringBuilder GetEnumScript()
    {
        StringBuilder enumScript = new StringBuilder();
        foreach (var typeInfo in _typeDescriptions)
        {
            if (typeInfo._values != null) // Need to improve to be compatible with futrue special type
            {
                enumScript.Append(string.Format("\n\tpublic enum {0} {{\n\t\t", typeInfo._typeName));
                enumScript.Append(typeInfo._values[0]);
                for (int i = 1; i < typeInfo._values.Length; ++i)
                {
                    enumScript.Append(",\n\t\t");
                    enumScript.Append(typeInfo._values[i]);
                }
                enumScript.Append("\n\t}");
            }
        }
        return enumScript;
    }

    /// <summary>
    /// Parse type to get the description of type
    /// </summary>
    /// <param name="typeName">Modified type name</param>
    /// <param name="fieldName">Filed name</param>
    /// <param name="typeInfo">Contians the original type name and values for special type</param>
    /// <returns>Desription data of type</returns>
    private TypeDescription ParseTypeInfo(string typeName, string fieldName, string[] typeInfo)
    {
        if (_typeDescriptions == null)
        {
            _typeDescriptions = new List<TypeDescription>();
        }

        TypeDescription enumType = new TypeDescription
        {
            _typeName = typeName,
            _fieldName = fieldName,
            _values = typeInfo.Length >= 2 ? typeInfo[1].Split(',') : null
        };

        if (typeInfo[0] == TableDataType.k_EnumType && enumType._values == null)
        {
            StopCompile(string.Format("Error in compiling sheet {0}: ill-formed enum {1}",
                _sheet.SheetName, typeName));
            return null;
        }
        _typeDescriptions.Add(enumType);
        return enumType;
    }

    private void StopCompile(string errMsg)
    {
        _isStop = true;
        _isSucess = false;
        _errMsg = errMsg;
    }

    /// <summary>
    /// Check if the field names can match with tables
    /// </summary>
    private bool CheckHeader(List<string> fieldNames)
    {
        return fieldNames.SequenceEqual(_tableCompiler.FieldNames);
    }

    /// <summary>
    /// Check if the field types can match with tables
    /// </summary>
    private bool CheckTypes(List<string> fieldTypes)
    {
        return fieldTypes.SequenceEqual(_tableCompiler.FieldTypes);
    }
}
