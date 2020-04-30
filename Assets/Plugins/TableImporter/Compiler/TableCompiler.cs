using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

/// <summary>
/// Class to convert config table to scriptableobject
/// </summary>
public class TableCompiler
{
    private IWorkbook _workbook;
    private string _filePath;
    private List<string> _fieldNames;
    private List<string> _fieldTypes;
    // Flag to indicate whether the compile process success
    private bool _isSucess = true;

    public string FilePath
    {
        get
        {
            return _filePath;
        }
    }

    public List<string> ErrorMessage { get; }

    public string EntityName { get; private set; }

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
        LoadWorkBook(path);
        _filePath = path;
        EntityName = Config.EntityPrefix + Path.GetFileNameWithoutExtension(path);
        ErrorMessage = new List<string>();
    }

    public TableCompiler(IWorkbook table, string path, string tableName)
    {
        _workbook = table;
        _filePath = path;
        EntityName = Config.EntityPrefix + Path.GetFileNameWithoutExtension(tableName);
        ErrorMessage = new List<string>();
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
        RecordCompileResult(sheetCompiler, true);
        if (!_isSucess)
        {
            return true;
        }
        return !_fieldNames.SequenceEqual(entityInfo.Attribute.FieldNames) ||
            !_fieldTypes.SequenceEqual(entityInfo.Attribute.FieldTypes);
    }

    private void LoadWorkBook(string path)
    {
        using (FileStream fileStream = 
            File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            _workbook = WorkbookFactory.Create(fileStream);
        }
    }   

    private void CompileSheet(int sheetIdx)
    {
        ISheet sheet = _workbook.GetSheetAt(sheetIdx);
        SheetCompiler sheetCompiler = new SheetCompiler(sheet, this);
        sheetCompiler.Compile();
        RecordCompileResult(sheetCompiler);
    }

    private void RecordCompileResult(SheetCompiler sheetCompiler, bool isSilenceMode = false)
    {
        _isSucess &= sheetCompiler.CompileSuccess;
        if (!_isSucess && !isSilenceMode)
        {
            ErrorMessage.Add(sheetCompiler.ErrorMessage);
        }
    }
}
