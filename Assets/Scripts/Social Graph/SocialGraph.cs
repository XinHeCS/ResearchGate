using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class SocialGraph
{
    // Total number of characters in this graph
    protected int _count;

    // The social graph
    protected Dictionary<Character, List<RelationDescription>> _graph =
        new Dictionary<Character, List<RelationDescription>>();

    #region Propertise

    public int Count
    {
        get
        {
            return _count;
        }
    }

    #endregion

    public void AddCharacter(Character ch)
    {
        if (!_graph.ContainsKey(ch))
        {
            _graph.Add(ch, new List<RelationDescription>());
            ++_count;
        }
    }

    public bool RemoveCharacter(Character ch)
    {
        var ret = _graph.Remove(ch);
        if (ret)
        {
            --_count;
        }
        return ret;
    }

    public bool HasCharacter(Character ch)
    {
        return _graph.ContainsKey(ch);
    }

    public Character GetCharacterByName(string name)
    {
        foreach (Character ch in _graph.Keys)
        {
            if (ch.Name == name)
            {
                return ch;
            }
        }
        return null;
    }

    public List<Character> GetCharactersByName(string name)
    {
        List<Character> retList = new List<Character>();

        foreach (var ch in _graph.Keys)
        {
            if (ch.Name == name)
            {
                retList.Add(ch);
            }
        }
        return retList;
    }

    public RelationDescription GetRelationDescription(Character sourceCh, Character targetCh)
    {
        if (_graph.TryGetValue(sourceCh, out List<RelationDescription> relations))
        {
            foreach (var relation in relations)
            {
                if (relation._targetCh.Equals(targetCh))
                {
                    return relation;
                }
            }
        }
        return null;
    }

    public List<RelationDescription> GetRelationDescriptions(Character character)
    {
        if (_graph.TryGetValue(character, out List<RelationDescription> relations))
        {
            return relations;
        }
        return null;
    }
}
