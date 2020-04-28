using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class Character : IEquatable<Character>
{
    public enum GenderType
    {
        male,
        female
    }

    protected string _md5;

    // The name of character
    protected string _name;

    // Gender of character
    protected GenderType _gender;

    // Birthday of character
    protected DateTime _birthday;

    // The importance of character in social graph
    protected int _honor;

    // Status of character
    protected BaseStatus _status;

    // History of character
    protected List<BaseSocialEvent> _history;

    #region Propertise

    public string MD5
    {
        get
        {
            return _md5;
        }
    }

    public string Name
    {
        get
        {
            return _name;
        }
    }

    public GenderType Gender
    {
        get
        {
            return _gender;
        }
    }

    public DateTime Birthday
    {
        get
        {
            return _birthday;
        }
    }

    public int Age
    {
        get
        {
            return DateTime.Now.Year - _birthday.Year;
        }
    }

    public int Honor
    {
        get
        {
            return _honor;
        }
        set
        {
            _honor = value;
        }
    }

    public BaseStatus Status
    {
        get
        {
            return _status;
        }
    }

    #endregion

    public void InsertEvent(BaseSocialEvent evt)
    {
        if (_history == null)
        {
            _history = new List<BaseSocialEvent>();
        }
        _history.Add(evt);
    }

    // Constructor
    public Character(string name, GenderType gender, DateTime birthday, int honor = 0)
    {
        _name = name;
        _gender = gender;
        _birthday = birthday;
        _honor = honor;

        _md5 = MD5Helper.GetMd5Hash(name + (int)_gender + _birthday.ToString());
    }

    public bool Equals(Character other)
    {
        return _md5 == other._md5;
    }
}
