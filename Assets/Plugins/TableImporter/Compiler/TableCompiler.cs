using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/// <summary>
/// Class to convert config table to scriptableobject
/// </summary>
public class TableCompiler
{
    public class TypeDescription
    {
        public string _typeName; // Type name
        public string _fieldName;
        public string[] _values; // Value domain for special type (e.g. enum)
    }

    private IWorkbook _workbook;
    private string _filePath;
    private List<string> _fieldNames;
    private List<string> _fieldTypes;
    private List<SheetCompiler> _sheetCompilers;
    private List<TypeDescription> _typeDescriptions; // For definition of enum field in sheet

    // Flag to indicate whether the compile process success
    private bool _isSucess = true;

    #region Propertise

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

    #endregion

    public TableCompiler(string path)
    {
        LoadWorkBook(path);
        _filePath = path;
        EntityName = Config.EntityPrefix + Path.GetFileNameWithoutExtension(path);
        ErrorMessage = new List<string>();
        InitSheetCompilers();
    }

    public void Compile()
    {
        foreach (var sheetCompiler in _sheetCompilers)
        {
            sheetCompiler.Compile();
            RecordCompileResult(sheetCompiler);
        }
        if (_isSucess)
        {
            GetTypeDescriptions();
            GenerateScript();
        }
    }

    public bool NeedCompile(TableImporter.TableEntityInfo entityInfo)
    {
        if (entityInfo == null)
        {
            return true;
        }

        foreach (var sheetCompiler in _sheetCompilers)
        {
            sheetCompiler.ParseSheet(); // Obnly parsing instead of compiling here
            RecordCompileResult(sheetCompiler, true);
            if (!_isSucess)
            {
                return true;
            }
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

    private void InitSheetCompilers()
    {
        _sheetCompilers = new List<SheetCompiler>();
        for (int i = 0; i < _workbook.NumberOfSheets; ++i)
        {
            _sheetCompilers.Add(new SheetCompiler(_workbook.GetSheetAt(i), this));
        }
    }

    private void RecordCompileResult(SheetCompiler sheetCompiler, bool isSilenceMode = false)
    {
        _isSucess &= sheetCompiler.CompileSuccess;
        if (!_isSucess && !isSilenceMode)
        {
            ErrorMessage.Add(sheetCompiler.ErrorMessage);
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
        sb.Replace(TemplateSymbol.k_ENTITYNAME, EntityName);
        // Fill Entity Body
        sb.Replace(TemplateSymbol.k_FIELDS, GetEntityBody().ToString());
        // Fill extra definitions
        sb.Replace(TemplateSymbol.k_ENUMDEF, GetEnumScript().ToString());

        // Fill table content in template
        var excelPath = FilePath;
        var assetName = EntityName.Replace(Config.EntityPrefix, Config.ScriptablePrefix);
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
        fieldNames.Append(FieldNames[0]);
        fieldNames.Append("\"");
        for (int i = 1; i < FieldNames.Count; ++i)
        {
            fieldNames.Append(", \"");
            fieldNames.Append(FieldNames[i]);
            fieldNames.Append("\"");
        }

        return fieldNames;
    }

    private StringBuilder GetFieldTypeInAttr()
    {
        StringBuilder fieldTypes = new StringBuilder("\"");
        fieldTypes.Append(FieldTypes[0]);
        fieldTypes.Append("\"");
        for (int i = 1; i < FieldTypes.Count; ++i)
        {
            fieldTypes.Append(", \"");
            fieldTypes.Append(FieldTypes[i]);
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

    private void GetTypeDescriptions()
    {
        for (int i = 0; i < FieldTypes.Count; ++i)
        {
            var typeInfo = FieldTypes[i].Split('|');
            var type = typeInfo[0];
            if (type == TableDataType.k_EnumType)
            {
                type += "_" + FieldNames[i]; // Modify type name for enum
            }
            ParseTypeInfo(type, FieldNames[i], typeInfo);
        }
    }

    /// <summary>
    /// Parse type to get the description of type
    /// </summary>
    /// <param name="typeName">Modified type name</param>
    /// <param name="fieldName">Filed name</param>
    /// <param name="typeInfo">Contians the original type name and values for special type</param>
    /// <returns>Desription data of type</returns>
    private void ParseTypeInfo(string typeName, string fieldName, string[] typeInfo)
    {
        if (_typeDescriptions == null)
        {
            _typeDescriptions = new List<TypeDescription>();
        }

        TypeDescription typeDec = new TypeDescription
        {
            _typeName = typeName,
            _fieldName = fieldName,
            _values = typeInfo.Length >= 2 ? typeInfo[1].Split(',') : null
        };

        if (typeInfo[0] == TableDataType.k_EnumType && typeDec._values == null)
        {
            _isSucess &= false;
            ErrorMessage.Add(string.Format("Error in {0}\\{1}: ill-formed enum {1}",
                FilePath, typeName));
        }
        _typeDescriptions.Add(typeDec);
    }
}
