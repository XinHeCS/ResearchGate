using System.Collections;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

/// <summary>
/// Class to convert config table to scriptableobject
/// </summary>
public class TableCompiler
{
    public const string k_IntType = "int";
    public const string k_BoolType = "bool";
    public const string k_FloatType = "float";
    public const string k_StringType = "string";
    public const string k_EnumType = "enum";

    private IWorkbook _workbook;
    private List<string> _fieldNames;
    private List<string> _fieldTypes;
    // Flag to indicate whether the compile process success
    private bool _isSucess = true;

    public List<string> ErrorMessage { get; }

    public bool CompileSuccess
    {
        get
        {
            return _isSucess;
        }
    }

    public List<string> FieldNames
    {
        get
        {
            return _fieldNames;
        }
        set
        {
            _fieldNames = value;
        }
    }

    public List<string> FieldTypes
    {
        get
        {
            return _fieldTypes;
        }
        set
        {
            _fieldTypes = value;
        }
    }

    public TableCompiler(string path)
    {
        _workbook = WorkbookFactory.Create(path);
    }

    public TableCompiler(IWorkbook table)
    {
        _workbook = table;
    }

    public void Compile()
    {
        for (int i = 0; i < _workbook.NumberOfSheets; ++i)
        {
            // Continue to compile event if some sheet has compile errors
            CompileSheet(i);
        }
    }

    public bool NeedCompile(TableImporter.TableEntityInfo entityInfo)
    {
        if (entityInfo == null)
        {
            return true;
        }
        
        SheetCompiler sheetCompiler = new SheetCompiler(_workbook.GetSheetAt(0), this);
        sheetCompiler.ParseSheet();
        RecordCompileResult(sheetCompiler);
        return _fieldNames.SequenceEqual(entityInfo.Attribute.FieldNames) &&
            _fieldTypes.SequenceEqual(entityInfo.Attribute.FieldTypes);
    }

    public bool HasType(string type)
    {
        switch (type)
        {
            case k_IntType:
                return true;
            case k_BoolType:
                return true;
            case k_FloatType:
                return true;
            case k_StringType:
                return true;
            case k_EnumType:
                return true;
            default:
                return false;
        }
    }

    private void CompileSheet(int sheetIdx)
    {
        ISheet sheet = _workbook.GetSheetAt(sheetIdx);
        SheetCompiler sheetCompiler = new SheetCompiler(sheet, this);
        sheetCompiler.Compile();
        RecordCompileResult(sheetCompiler);
    }

    private void RecordCompileResult(SheetCompiler sheetCompiler)
    {
        _isSucess &= sheetCompiler.CompileSuccess;
        if (!_isSucess)
        {
            ErrorMessage.Add(sheetCompiler.ErrorMessage);
        }
    }
}
