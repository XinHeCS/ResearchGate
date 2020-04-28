using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class SocialEventTools
{
    public static BaseSocialEvent CreateEvent(int type)
    {
        switch (type)
        {
            case SocialEventType.k_Pick:
                return null;
            case SocialEventType.k_Test:
                return null;
            default:
                return null;
        }
    }
}
