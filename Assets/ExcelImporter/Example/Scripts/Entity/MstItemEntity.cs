using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System;

[Serializable]
[TableEntity(FieldNames = new string[] { "a", "b"}, FieldTypes = new string[] { "int", "float"})]
public class MstItemEntity
{
	public int id;
	public string name;
	public int price;
	public bool isNotForSale;
	public float rate;
	public MstItemCategory category;
}

public enum MstItemCategory
{
	Red,
	Green,
	Blue,
}