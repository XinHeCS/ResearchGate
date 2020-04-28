using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using UnityEngine;

public class SheetCompiler
{
    class EnumDefinition
    {
        public string _name;
        public string[] _values;
    }

    private List<EnumDefinition> _enumDefs; // For definition of enum field in sheet
    private TableCompiler _tableCompiler;
    private ISheet _sheet;

    // Flag to control whether stop the compiling process
    private bool _isStop = false;
    // Flag to indicate whether the compile process success
    private bool _isSucess = false;

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
                _isStop = true;
                _errMsg = string.Format("Error in compiling sheet {0} ( {1} : {2} ): Unknown header is not allowed",
                    _sheet.SheetName, cell.RowIndex, i);
                return;
            }
        }
        if (_tableCompiler.FieldNames == null)
        {
            _tableCompiler.FieldNames = fieldNames;
        }
        else if (!CheckHeader(fieldNames))
        {
            _isStop = true;
            _errMsg = string.Format("Error in compiling sheet {0}: Unmatched header fields", _sheet.SheetName);
        }
    }

    private void ParseFieldType()
    {
        List<string> fieldTypes = new List<string>();
        IRow typeRow = _sheet.GetRow(1);
        if (typeRow == null)
        {
            _isStop = true;
            _isSucess = true;
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
                var type = cell.StringCellValue.Trim().ToLower();
                if (!_tableCompiler.HasType(type))
                {
                    _isStop = true;
                    _errMsg = string.Format("Error in compiling sheet {0} ( {1} : {2} ): Unsuportted type {3}",
                        _sheet.SheetName, cell.RowIndex, i, type);
                    return;
                }
                fieldTypes.Add(type);
                if (type == TableCompiler.k_EnumType)
                {
                    GetEnumDefinition(i);
                }
            }
            else if (cell.CellType == CellType.Unknown)
            {
                _isStop = true;
                _errMsg = string.Format("Error in compiling sheet {0} ( {1} : {2} ): Unknown field is not allowed",
                    _sheet.SheetName, cell.RowIndex, i);
                return;
            }
        }
        if (_tableCompiler.FieldTypes == null)
        {
            _tableCompiler.FieldTypes = fieldTypes;
        }
        else if (!CheckTypes(fieldTypes))
        {
            _isStop = true;
            _errMsg = string.Format("Error in compiling sheet {0}: Unmatched field type", _sheet.SheetName);
        }
    }

    private void GenerateScript()
    {
        var templatepath = Path.Combine(new string[] { Directory.GetCurrentDirectory(), "Template" });
       
        StringBuilder sb = new StringBuilder();
        sb.Append(File.ReadAllText(templatepath));

        // Fill Field names in attribute
        sb.Replace(TableCompiler.k_FIELDNAMES, GetFieldNameInAttr().ToString());
        // Fill Field types in attribute
        sb.Replace(TableCompiler.k_FIELDTYPES, GetFieldTypeInAttr().ToString());
        // Fill Entity names
        sb.Replace(TableCompiler.k_ENTITYNAME, _tableCompiler.EntityName);
        // Fill Entity Body
        sb.Replace(TableCompiler.k_FIELDS, GetEntityBody().ToString());
        // Fill extra definitions
        sb.Append(GetExtraScript().ToString());

        // Write to disk
        File.WriteAllText(
            Path.Combine(TableCompiler.TableScriptablePath, _tableCompiler.EntityName + ".cs"),
            sb.ToString()
            );
        _isSucess = true;
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
        StringBuilder body = new StringBuilder("public ");
        body.Append(_tableCompiler.FieldTypes[0]);
        body.Append(" ");
        body.Append(_tableCompiler.FieldNames[0]);
        body.Append(";");
        for (int i = 1; i < _tableCompiler.FieldTypes.Count; ++i)
        {
            body.Append("\npublic ");
            body.Append(_tableCompiler.FieldTypes[i]);
            body.Append(" ");
            body.Append(_tableCompiler.FieldNames[i]);
            body.Append(";");
        }

        return body;
    }

    private StringBuilder GetExtraScript()
    {
        StringBuilder extra = new StringBuilder();
        foreach (var enumDef in _enumDefs)
        {
            extra.Append(string.Format("\npublic enum {0} {\n\t", enumDef._name));
            extra.Append(enumDef._values[0]);
            for (int i = 1; i < enumDef._values.Length; ++i)
            {
                extra.Append(",\n\t");
                extra.Append(enumDef._values[i]);
            }
            extra.Append("\n}");
        }
        return extra;
    }

    private void GetEnumDefinition(int col)
    {
        var validations = _sheet.GetDataValidations();
        // Find matched validation of current column
        foreach (var validation in validations)
        {
            CellRangeAddressList addressRange = validation.Regions;
            foreach (var range in addressRange.CellRangeAddresses)
            {
                if (range.FirstColumn <= col && col <= range.LastColumn)
                {
                    EnumDefinition def = new EnumDefinition
                    {
                        _name = _tableCompiler.FieldNames[col],
                        _values = validation.ValidationConstraint.ExplicitListValues
                    };
                    _enumDefs.Add(def);
                    return;
                }
            }
        }
        _isStop = true;
        _errMsg = string.Format("Error in compiling sheet {0}: Can't find definition of enum {1}",
            _sheet.SheetName, _tableCompiler.FieldNames[col]);
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
