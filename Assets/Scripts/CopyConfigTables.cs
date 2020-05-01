using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using UnityEditor;

public class CopyConfigTables 
{
    private const string _sheetPath = @"Assets\ConfigTable\Test";
    private static int _testCount = 0;

    [MenuItem("Assets/Create/TestTables", false)]
    static void CreateTables()
    {
        IWorkbook workbook = WorkbookFactory.Create(_sheetPath + ".xlsx");
        IWorkbook newBook = new XSSFWorkbook();


        for (int i = 0; i < 20; ++i)
        {
            File.Copy(_sheetPath + ".xlsx", string.Format(_sheetPath + "_{0}.xlsx", _testCount++), true);
        }

        //newBook.Write(new FileStream(string.Format(_sheetPath + "_{0}.xlsx", _testCount), FileMode.OpenOrCreate));
    }
}
