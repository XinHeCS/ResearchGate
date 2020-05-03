using System.Collections;
using System.Collections.Generic;

public class TableDataType
{
    public const string k_IntType = "int";
    public const string k_BoolType = "bool";
    public const string k_FloatType = "float";
    public const string k_StringType = "string";
    public const string k_EnumType = "enum";

    public static bool HasType(string type)
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
}
