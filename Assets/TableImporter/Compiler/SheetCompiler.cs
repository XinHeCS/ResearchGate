using System.Collections;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

public class SheetCompiler
{
    private string _script;
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
        // TODO ...
    }

    public void ParseSheet()
    {
        ParseFieldName();
        if (_isStop)
        {
            return;
        }
        ParseFieldType();
        if (_isStop)
        {
            return;
        }
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
