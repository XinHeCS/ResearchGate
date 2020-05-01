using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class BaseStatus
{
    // academic level
    public int _acl;

    // engineer level
    public int _egl;

    // social level
    public int _scl;

    public override string ToString()
    {
        return string.Format("Academic : {0}, Engineer : {1}, Social : {2}", _acl, _egl, _scl);
    }
}
