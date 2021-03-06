﻿using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using UnityEditor.Compilation;
using System;
using System.IO;
using System.Reflection;
using NPOI.SS.UserModel;

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
                return string.IsNullOrEmpty(Attribute.ExcelPath) ? AssetType.Name : Attribute.ExcelPath;
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

    static TableImporter()
    {
        // Register re-compile event for importer since Unity will reload all 
        // assemblise after compiling.
        // To prevent the clear event behaviour, we need to regidter finish compiling event
        // every time the impoter is re-loaded.
        //CompilationPipeline.assemblyCompilationFinished += OnScriptsFinishCompiled;
        AssemblyReloadEvents.afterAssemblyReload += OnAssemblyLoad;
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
        bool hasDoneCompile = false;

        foreach (string path in importedAssets)
        {
            if (Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx")
            {
                var excelName = Path.GetFileNameWithoutExtension(path);
                if (excelName.StartsWith("~$")) continue;

                if (!CompileTable(path, out hasDoneCompile))
                {
                    continue;
                }
                imported = true;
            }
        }

        if (imported)
        {
            if (!hasDoneCompile)
            {
                // No file has been re-compiled in this import process
                // Just read all table data. 
                UpdateImportedTable(importedAssets);
            }
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }
    }

    //static void OnScriptsFinishCompiled(string path, CompilerMessage[] messages)
    //{
    //    var asmName = string.IsNullOrEmpty(Config.TableAssemblyName) ?
    //        Config.DefaultAssembly : Config.TableAssemblyName;
    //    if (asmName == Path.GetFileNameWithoutExtension(path))
    //    {
    //        Directory.SetLastWriteTime(Config.TableScriptablePath, DateTime.Now);
    //    }
    //}

    static void OnAssemblyLoad()
    {
        if (Directory.Exists(Config.TableScriptablePath))
        {
            // After compile process, we need to flush the cached data
            FlushTableAttrInfo();

            var assets = Directory.GetFiles(Config.TableScriptablePath, string.Format("{0}*.cs", Config.ScriptablePrefix));
            foreach (var asset in assets)
            {
                if (Directory.GetCreationTime(asset) !=
                    Directory.GetLastWriteTime(asset))
                {
                    var assetInfo = cachedAssetInfos.Find(i => i.AssetType.Name == Path.GetFileNameWithoutExtension(asset));
                    if (assetInfo != null)
                    {
                        LoadTableData(assetInfo);
                        Directory.SetCreationTime(asset, Directory.GetLastWriteTime(asset));
                    }
                }
            }
        }
    }

    /// <summary>
    /// Reload all re-compiled table data
    /// </summary>
    static void UpdateCompiledTable()
    {
        foreach (var assetInfo in cachedAssetInfos)
        {
            LoadTableData(assetInfo);
        }
    }

    /// <summary>
    /// Reload the imported table data
    /// </summary>
    /// <param name="excelPath">Imported files list</param>
    static void UpdateImportedTable(string[] excelPath)
    {
        foreach (var path in excelPath)
        {
            var assetInfo = cachedAssetInfos.Find(i => i.Attribute.ExcelPath == path);
            if (assetInfo != null)
            {
                LoadTableData(assetInfo);
            }
        }
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
                FileUtil.DeleteFileOrDirectory(path + ".meta");
                if (excelName.StartsWith("~$"))
                {
                    continue;
                }
                var assetPath =
                    Path.Combine(Config.TableDataPath, string.Format("{0}{1}.asset", Config.ScriptablePrefix, excelName));
                FileUtil.DeleteFileOrDirectory(assetPath);
                FileUtil.DeleteFileOrDirectory(assetPath + ".meta");
                var scriptablePath =
                    Path.Combine(Config.TableScriptablePath, string.Format("{0}{1}.cs", Config.ScriptablePrefix, excelName));
                FileUtil.DeleteFileOrDirectory(scriptablePath);
                FileUtil.DeleteFileOrDirectory(assetPath + ".meta");
            }
        }
    }

    static bool CompileTable(string path, out bool hasDoneCompile)
    {
        hasDoneCompile = false;
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
            hasDoneCompile = true;
        }
        return true;
    }

    /// <summary>
    /// This fucntion will be called the first time tables are imported
    /// </summary>
    static void FlushTableAttrInfo(System.Reflection.Assembly assembly = null)
    {
        // Clear dirty cache
        cachedEntityInfos = new List<TableEntityInfo>();
        cachedAssetInfos = new List<TableAssetInfo>();
        if (assembly == null)
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                var tableAsm = string.IsNullOrEmpty(Config.TableAssemblyName) ? 
                    Config.DefaultAssembly : Config.TableAssemblyName;
                if (asm.GetName().Name == tableAsm)
                {
                    GetTableAttrInfo(asm);
                }
            }
        }
        else
        {
            GetTableAttrInfo(assembly);
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
            return WorkbookFactory.Create(stream);
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
                if (fieldInfo.FieldType.IsEnum) return Enum.Parse(fieldInfo.FieldType, cell.StringCellValue.ToLower());
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

    static void GetEntityListFromSheet(ISheet sheet, Type entityType, object list, MethodInfo addMethod)
    {
        List<string> excelColumnNames = GetFieldNamesFromSheetHeader(sheet);

        // row of index 0 is header and 1 is type definition
        for (int i = 2; i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);
            if (row == null) break;

            ICell entryCell = row.GetCell(0);
            if (entryCell == null || entryCell.CellType == CellType.Blank) break;

            // skip comment row
            if (entryCell.CellType == CellType.String && entryCell.StringCellValue.StartsWith("#")) continue;

            var entity = CreateEntityFromRow(row, excelColumnNames, entityType, sheet.SheetName);
            addMethod.Invoke(list, new object[] { entity });
        }
    }

    static void LoadTableData(TableAssetInfo info)
    {
        string assetPath = "";
        string assetName = info.AssetType.Name + ".asset";

        assetPath = Path.Combine(info.Attribute.AssetPath, assetName);
        
        UnityEngine.Object asset = LoadOrCreateAsset(assetPath, info.AssetType);

        var assetFields = info.AssetType.GetFields(BindingFlags.NonPublic | BindingFlags.Instance);

        IWorkbook book = LoadBook(info.Attribute.ExcelPath);

        // Get filed and type infomation
        Type fieldType = assetFields[0].FieldType;
        if (!fieldType.IsGenericType || (fieldType.GetGenericTypeDefinition() != typeof(List<>))) return;
        Type entityType = fieldType.GetGenericArguments()[0];
        Type listType = typeof(List<>).MakeGenericType(entityType);
        MethodInfo listAddMethod = listType.GetMethod("Add", new Type[] { entityType });

        // Read data from table
        object list = Activator.CreateInstance(listType);
        for (int i = 0; i < book.NumberOfSheets; ++i)
        {
            ISheet sheet = book.GetSheetAt(i);
            if (sheet == null) continue;

            GetEntityListFromSheet(sheet, entityType, list, listAddMethod);
        }
        assetFields[0].SetValue(asset, list);

        if (info.Attribute.LogOnImport)
        {
            Debug.Log(string.Format("Imported {0} sheets form {1}.", book.NumberOfSheets, info.Attribute.ExcelPath));
        }

        EditorUtility.SetDirty(asset);
    }
}
