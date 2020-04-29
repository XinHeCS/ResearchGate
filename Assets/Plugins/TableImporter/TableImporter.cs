using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using UnityEditor.Compilation;

public class TableImporter : AssetPostprocessor
{
    public class TableAssetInfo
    {
        public Type AssetType { get; set; }
        public TableAssetAttribute Attribute { get; set; }
        public string TableName
        {
            get
            {
                return string.IsNullOrEmpty(Attribute.ExcelName) ? AssetType.Name : Attribute.ExcelName;
            }
        }
    }

    public class TableEntityInfo
    {
        public Type EntityType { get; set; }
        public TableEntityAttribute Attribute { get; set; }
        public string EntityName
        {
            get
            {
                return EntityType.Name;
            }
        }
    }

    static List<TableAssetInfo> cachedAssetInfos = null;
    static List<TableEntityInfo> cachedEntityInfos = null; //  Clear on compile.

    // Register compile event
    static TableImporter()
    {
        CompilationPipeline.assemblyCompilationFinished += OnScriptsFinishCompiled;
    }

    //void OnPreprocessAsset()
    //{
    //    var path = assetImporter.assetPath;
    //    if (Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx")
    //    {
    //        var excelName = Path.GetFileNameWithoutExtension(path);
    //        if (excelName.StartsWith("~$"))
    //        {
    //            return;
    //        }
    //        if (cachedEntityInfo == null)
    //        {
    //            cachedEntityInfo = FindTableEntityInfo();
    //        }
    //        var entityInfo = cachedEntityInfo.Find(
    //            i => i.EntityName.Equals(Config.EntityPrefix + excelName)
    //            );
    //        TableCompiler compiler = new TableCompiler(path);
    //        if (compiler.NeedCompile(entityInfo))
    //        {
    //            compiler.Compile();
    //            if (!compiler.CompileSuccess)
    //            {
    //                foreach (var msg in compiler.ErrorMessage)
    //                {
    //                    Debug.LogError(msg);
    //                }
    //                return;
    //            }
    //            Debug.Log(string.Format("Cmopiled table {0}", excelName));
    //        }
    //    }
    //}

    static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
    {
        HandleImportFile(importedAssets);
        HandleDeletedFile(deletedAssets);
    }

    static void HandleImportFile(string[] importedAssets)
    {
        bool imported = false;
        foreach (string path in importedAssets)
        {
            if (Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx")
            {
                var excelName = Path.GetFileNameWithoutExtension(path);
                if (excelName.StartsWith("~$")) continue;

                if (!CompileTable(path))
                {
                    continue;
                }
                imported = true;
            }
        }

        if (imported)
        {
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }
    }

    static void OnScriptsFinishCompiled(string assemblyPath, CompilerMessage[] msg)
    {
        foreach (var ms in msg)
        {
            if (ms.type == CompilerMessageType.Error)
            {
                return;
            }
        }
        FlushTableAttrInfo();
    }

    static void GetTableAttrInfo(System.Reflection.Assembly assembly)
    {
        foreach (var type in assembly.GetTypes())
        {
            var entityAttrs = type.GetCustomAttributes<TableEntityAttribute>(false);
            foreach (var attr in entityAttrs)
            {
                var entityInfo = new TableEntityInfo
                {
                    EntityType = type,
                    Attribute = attr
                };
                cachedEntityInfos.Add(entityInfo);
            }
            var assetAttrs = type.GetCustomAttributes<TableAssetAttribute>(false);
            foreach (var attr in assetAttrs)
            {
                var assetInfo = new TableAssetInfo
                {
                    AssetType = type,
                    Attribute = attr
                };
                cachedAssetInfos.Add(assetInfo);
            }
        }
    }

    static void HandleDeletedFile(string[] deletedAssets)
    {
        foreach (string path in deletedAssets)
        {
            if (Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx")
            {
                var excelName = Path.GetFileNameWithoutExtension(path);
                if (excelName.StartsWith("~$"))
                {
                    FileUtil.DeleteFileOrDirectory(path + ".meta");
                    AssetDatabase.Refresh();
                }
            }
        }
    }

    static bool CompileTable(string path)
    {
        var excelName = Path.GetFileNameWithoutExtension(path);

        if (cachedEntityInfos == null)
        {
            FlushTableAttrInfo();
        }
        var entityInfo = cachedEntityInfos.Find(
            i => i.EntityName.Equals(Config.EntityPrefix + excelName)
            );
        TableCompiler compiler = new TableCompiler(path);
        if (compiler.NeedCompile(entityInfo))
        {
            compiler.Compile();
            if (!compiler.CompileSuccess)
            {
                foreach (var msg in compiler.ErrorMessage)
                {
                    Debug.LogError(msg);
                }
                return false;
            }
            Debug.Log(string.Format("Cmopiled table {0}", excelName));
        }
        return true;
    }

    /// <summary>
    /// This fucntion will be called the first time tables are imported
    /// </summary>
    static void FlushTableAttrInfo()
    {
        // Clear dirty cache
        cachedEntityInfos = new List<TableEntityInfo>();
        cachedAssetInfos = new List<TableAssetInfo>();
        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
            var asmName = assembly.GetName();
            if (string.IsNullOrEmpty(Config.TableAssemblyName) || asmName.Name == Config.TableAssemblyName)
            {
                GetTableAttrInfo(assembly);
            }
        }
    }

    static UnityEngine.Object LoadOrCreateAsset(string assetPath, Type assetType)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(assetPath));

        var asset = AssetDatabase.LoadAssetAtPath(assetPath, assetType);

        if (asset == null)
        {
            asset = ScriptableObject.CreateInstance(assetType.Name);
            AssetDatabase.CreateAsset((ScriptableObject)asset, assetPath);
            asset.hideFlags = HideFlags.NotEditable;
        }

        return asset;
    }

    static IWorkbook LoadBook(string excelPath)
    {
        using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            if (Path.GetExtension(excelPath) == ".xls") return new HSSFWorkbook(stream);
            else return new XSSFWorkbook(stream);
        }
    }

    static List<string> GetFieldNamesFromSheetHeader(ISheet sheet)
    {
        IRow headerRow = sheet.GetRow(0);

        var fieldNames = new List<string>();
        for (int i = 0; i < headerRow.LastCellNum; i++)
        {
            var cell = headerRow.GetCell(i);
            if (cell == null || cell.CellType == CellType.Blank) break;
            fieldNames.Add(cell.StringCellValue);
        }
        return fieldNames;
    }

    static object CellToFieldObject(ICell cell, FieldInfo fieldInfo, bool isFormulaEvalute = false)
    {
        var type = isFormulaEvalute ? cell.CachedFormulaResultType : cell.CellType;

        switch (type)
        {
            case CellType.String:
                if (fieldInfo.FieldType.IsEnum) return Enum.Parse(fieldInfo.FieldType, cell.StringCellValue);
                else return cell.StringCellValue;
            case CellType.Boolean:
                return cell.BooleanCellValue;
            case CellType.Numeric:
                return Convert.ChangeType(cell.NumericCellValue, fieldInfo.FieldType);
            case CellType.Formula:
                if (isFormulaEvalute) return null;
                return CellToFieldObject(cell, fieldInfo, true);
            default:
                if (fieldInfo.FieldType.IsValueType)
                {
                    return Activator.CreateInstance(fieldInfo.FieldType);
                }
                return null;
        }
    }

    static object CreateEntityFromRow(IRow row, List<string> columnNames, Type entityType, string sheetName)
    {
        var entity = Activator.CreateInstance(entityType);

        for (int i = 0; i < columnNames.Count; i++)
        {
            FieldInfo entityField = entityType.GetField(
                columnNames[i],
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
            );
            if (entityField == null) continue;
            if (!entityField.IsPublic && entityField.GetCustomAttributes(typeof(SerializeField), false).Length == 0) continue;

            ICell cell = row.GetCell(i);
            if (cell == null) continue;

            try
            {
                object fieldValue = CellToFieldObject(cell, entityField);
                entityField.SetValue(entity, fieldValue);
            }
            catch
            {
                throw new Exception(string.Format("Invalid excel cell type at row {0}, column {1}, {2} sheet.", row.RowNum, cell.ColumnIndex, sheetName));
            }
        }
        return entity;
    }

    static object GetEntityListFromSheet(ISheet sheet, Type entityType)
    {
        List<string> excelColumnNames = GetFieldNamesFromSheetHeader(sheet);

        Type listType = typeof(List<>).MakeGenericType(entityType);
        MethodInfo listAddMethod = listType.GetMethod("Add", new Type[] { entityType });
        object list = Activator.CreateInstance(listType);

        // row of index 0 is header
        for (int i = 1; i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);
            if (row == null) break;

            ICell entryCell = row.GetCell(0);
            if (entryCell == null || entryCell.CellType == CellType.Blank) break;

            // skip comment row
            if (entryCell.CellType == CellType.String && entryCell.StringCellValue.StartsWith("#")) continue;

            var entity = CreateEntityFromRow(row, excelColumnNames, entityType, sheet.SheetName);
            listAddMethod.Invoke(list, new object[] { entity });
        }
        return list;
    }

    static void ImportExcel(string excelPath, TableAssetInfo info)
    {
        string assetPath = "";
        string assetName = info.AssetType.Name + ".asset";

        if (string.IsNullOrEmpty(info.Attribute.AssetPath))
        {
            string basePath = Path.GetDirectoryName(excelPath);
            assetPath = Path.Combine(basePath, assetName);
        }
        else
        {
            var path = Path.Combine("Assets", info.Attribute.AssetPath);
            assetPath = Path.Combine(path, assetName);
        }
        UnityEngine.Object asset = LoadOrCreateAsset(assetPath, info.AssetType);

        IWorkbook book = LoadBook(excelPath);

        var assetFields = info.AssetType.GetFields();
        int sheetCount = 0;

        foreach (var assetField in assetFields)
        {
            ISheet sheet = book.GetSheet(assetField.Name);
            if (sheet == null) continue;

            Type fieldType = assetField.FieldType;
            if (!fieldType.IsGenericType || (fieldType.GetGenericTypeDefinition() != typeof(List<>))) continue;

            Type[] types = fieldType.GetGenericArguments();
            Type entityType = types[0];

            object entities = GetEntityListFromSheet(sheet, entityType);
            assetField.SetValue(asset, entities);
            sheetCount++;
        }

        if (info.Attribute.LogOnImport)
        {
            Debug.Log(string.Format("Imported {0} sheets form {1}.", sheetCount, excelPath));
        }

        EditorUtility.SetDirty(asset);
    }
}
