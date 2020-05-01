// ======================================================
//		Auto-generated code, don't modify it!
// ======================================================
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[Serializable]
[TableEntity(FieldNames = new string[] { "id", "name", "price", "isTest", "factor", "color" }, FieldTypes = new string[] { "int", "string", "float", "int", "float", "enum|red,green,blue" })]
public class Entity_Test_9
{
	public int id;
	public string name;
	public float price;
	public int isTest;
	public float factor;
	public enum_color color;

	public enum enum_color {
		red,
		green,
		blue
	}
}

[TableAsset(AssetPath = Config.TableDataPath, ExcelPath = "Assets/ConfigTable/Test_9.xlsx", LogOnImport = Config.LogOnImport)]
public class Scriptable_Test_9 : ScriptableObject
{
	[SerializeField]
	List<Entity_Test_9> _entitise = new List<Entity_Test_9>();

    public Entity_Test_9 this[int index] 
	{
		get 
		{
			if (0 <= index && index <= _entitise.Count)
            {
                return _entitise[index];
            }
            return null;
		}
	}

	public int Count
    {
        get
        {
            return _entitise.Count;
        }
    }
}
