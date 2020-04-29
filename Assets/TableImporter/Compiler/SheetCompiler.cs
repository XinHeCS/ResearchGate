using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

public class SheetCompiler
{
    class EnumType
    {
        public string _name;
        public string[] _values;
    }

    private List<EnumType> _enumTypes; // For definition of enum field in sheet
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
                fieldNames.Add(cell.StringCellValue.Trim().ToLower());
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
                    StopCompile(string.Format("Error in compiling sheet {0} ( {1} : {2} ): Unsuportted type {3}",
                        _sheet.SheetName, cell.RowIndex, i, type));
                    return;
                }
                if (type == TableDataType.k_EnumType)
                {
                    if (defInfo.Length < 2)
                    {
                        StopCompile(
                            string.Format("Error in compiling sheet {0}: Definition of enum type {1} is missing", _sheet.SheetName, type));
                    }
                    type += "_" + _tableCompiler.FieldNames[i];
                    GetEnumDefinition(type, defInfo[1]);
                }
                fieldTypes.Add(type);
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
        var templatepath = Path.Combine(
            new string[] 
            {
                Directory.GetCurrentDirectory(),
                Config.EntityTemplatePath
            });
       
        StringBuilder sb = new StringBuilder();
        sb.Append(File.ReadAllText(templatepath));

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

        var path = Path.Combine(
            new string[]
            {
                Directory.GetCurrentDirectory(),
                Config.TableEntityPath,
                _tableCompiler.EntityName + ".cs"
            });
        // Write to disk
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
        for (int i = 0; i < _tableCompiler.FieldTypes.Count; ++i)
        {
            if (i == 0)
            {
                body.Append("\tpublic ");
            }
            else
            {
                body.Append("\n\tpublic ");
            }
            body.Append(_tableCompiler.FieldTypes[i]);
            body.Append(" ");
            body.Append(_tableCompiler.FieldNames[i]);
            body.Append(";");
        }

        return body;
    }

    private StringBuilder GetEnumScript()
    {
        StringBuilder enumScript = new StringBuilder();
        foreach (var enumDef in _enumTypes)
        {
            enumScript.Append(string.Format("\n\tpublic enum {0} {{\n\t\t", enumDef._name));
            enumScript.Append(enumDef._values[0]);
            for (int i = 1; i < enumDef._values.Length; ++i)
            {
                enumScript.Append(",\n\t\t");
                enumScript.Append(enumDef._values[i]);
            }
            enumScript.Append("\n\t}");
        }
        return enumScript;
    }

    private EnumType GetEnumDefinition(string enumName, string enumDef)
    {
        if (_enumTypes == null)
        {
            _enumTypes = new List<EnumType>();
        }

        EnumType enumType = new EnumType
        {
            _name = enumName,
            _values = enumDef.Split(',')
        };

        if (enumType._values.Length < 2)
        {
            StopCompile(string.Format("Error in compiling sheet {0}: ill-formed enum {1}",
                _sheet.SheetName, enumName));
            return null;
        }
        _enumTypes.Add(enumType);
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
