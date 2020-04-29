// ======================================================
//		Auto-generated code, don't modify it!
// ======================================================
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[Serializable]
[TableEntity(FieldNames = new string[] { "id", "name", "price", "isTest", "rate", "color" }, FieldTypes = new string[] { "int", "string", "float", "bool", "float", "enum|red,green,blue" })]
public class Entity_Test
{
	public int id;
	public string name;
	public float price;
	public bool isTest;
	public float rate;
	public enum_color color;

	public enum enum_color {
		red,
		green,
		blue
	}
}

[TableAsset(AssetPath = Config.TableScriptablePath, ExcelName = "Test", LogOnImport = Config.LogOnImport)]
public class Scriptable_Test : ScriptableObject
{
	List<Entity_Test> _entitise = new List<Entity_Test>();

    public Entity_Test this[int index] 
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
