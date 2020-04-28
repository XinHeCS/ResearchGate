using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using NPOI.SS.UserModel;

/// <summary>
/// Class to convert config table to scriptableobject
/// </summary>
public class TableCompiler
{
    private string _script;
    private string[] _fieldNames;
    private string[] _fieldTypes;
    private List<TableEntityAttribute> _tmpInfo;

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

    public void Compile(string path, ref List<TableEntityAttribute> entityInfo)
    {
        _tmpInfo = entityInfo;

        IWorkbook workbook = WorkbookFactory.Create(path);
        for (int i = 0; i < workbook.NumberOfSheets; ++i)
        {
            ISheet sheet = workbook.GetSheetAt(i);
            CompileSheet(sheet);
            if (_isStop)
            {
                break;
            }
        }
    }

    public void CompileSheet(ISheet sheet)
    {
        var excelName = sheet.SheetName;
        IRow headerRow = sheet.GetRow(0);
        if (headerRow == null)
        {
            _isStop = true;
            _isSucess = true;
            return;
        }
        ParseFieldType(headerRow);

    }

    private void ParseFieldType(IRow headerRow)
    {
        _fieldNames = new string[headerRow.LastCellNum + 1];
        foreach (var cell in headerRow.Cells)
        {
            if (cell == null || cell.CellType == CellType.Blank)
            {
                _isStop = true;
                _errMsg = string.Format("Error in {0} : {1}: Blamk header is not allowed",
                    cell.RowIndex, cell.ColumnIndex);
                return;
            }
            if (cell.CellType == CellType.String)
            {
                _fieldNames[cell.ColumnIndex] = cell.StringCellValue.Trim();
            }
        }
    }
}
