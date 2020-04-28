using System.Collections;
using System.Collections.Generic;
using UnityEngine;

/// <summary>
/// Class to describe the relationships to another characters
/// </summary>
public class RelationDescription
{
    // The other character
    public Character _targetCh;

    // The relation type to the other character
    public RelationType _relationType;

    // Like factor to taregt character
    public int _likeFactor;

    // Dislike factor to target character
    public int _dislikeFactor;
}
