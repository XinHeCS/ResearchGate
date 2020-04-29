using System;
using System.Collections;
using System.Collections.Generic;

[Serializable]
[TableEntity(FieldNames = new string[] { "id", "name", "price", "istest", "rate", "color" }, FieldTypes = new string[] { "int", "string", "float", "bool", "float", "enum_color" })]
public class Entity_Test
{
	public int id;
	public string name;
	public float price;
	public bool istest;
	public float rate;
	public enum_color color;

	public enum enum_color {
		red,
		green,
		blue
	}
}
