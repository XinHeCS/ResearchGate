using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

public class SheetCompiler
{
    protected const string k_ErrorMsgHeader = "Error in {0}\\{1} (Row: {2}, Column: {3})";
    protected const string k_SimpleErrMsgHeader = "Error in {0}\\{1}";    

    private TableCompiler _tableCompiler;
    private ISheet _sheet;

    // Flag to indicate whether the compile process success
    private bool _isSucess = true;
    // Flag to indicate whether the compile has parseed all sheets before
    private bool _hasParsed = false;

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
        if (!_hasParsed)
        {
            ParseSheet();
        }
    }

    public void ParseSheet()
    {
        ParseFieldName();
        ParseFieldType();
        _hasParsed = true;
    }

    private void ParseFieldName()
    {
        List<string> fieldNames = new List<string>();
        IRow headerRow = _sheet.GetRow(0);
        if (headerRow == null)
        {
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
                StopCompile(
                    string.Format(k_ErrorMsgHeader + ": Filed name should be a string.",
                    _tableCompiler.FilePath,
                    _sheet.SheetName, cell.RowIndex + 1, i + 1));
                return;
            }
        }
        if (_tableCompiler.FieldNames == null)
        {
            _tableCompiler.FieldNames = fieldNames;
        }
        else if (!CheckHeader(fieldNames))
        {
            StopCompile(string.Format(k_SimpleErrMsgHeader + ": Field name in different sheets should keep identical.", 
                _tableCompiler.FilePath, _sheet.SheetName));
        }
    }

    private void ParseFieldType()
    {
        List<string> fieldTypes = new List<string>();
        IRow typeRow = _sheet.GetRow(1);
        if (typeRow == null)
        {
            StopCompile(
                string.Format(k_SimpleErrMsgHeader + ": Missing type definitoins.", 
                _tableCompiler.FilePath, _sheet.SheetName));
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
                    StopCompile(string.Format(k_ErrorMsgHeader + ": (3) type is unsupported",
                        _tableCompiler.FilePath, _sheet.SheetName, cell.RowIndex + 1, i + 1, type));
                    return;
                }
                fieldTypes.Add(rawTypeInfo);
            }
            else if (cell.CellType == CellType.Unknown)
            {
                StopCompile(string.Format(k_ErrorMsgHeader + ": Unrecognizable type",
                    _tableCompiler.FilePath, _sheet.SheetName, cell.RowIndex + 1, i + 1));
                return;
            }
        }
        if (_tableCompiler.FieldTypes == null)
        {
            _tableCompiler.FieldTypes = fieldTypes;
        }
        else if (!CheckTypes(fieldTypes))
        {
            StopCompile(string.Format(k_SimpleErrMsgHeader + ": type definition in different sheets should keep identical.", 
                _tableCompiler.FilePath, _sheet.SheetName));
        }
    }


    private void StopCompile(string errMsg)
    {
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
