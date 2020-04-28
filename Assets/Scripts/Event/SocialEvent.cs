using System.Collections;
using System.Collections.Generic;
using UnityEngine;

/// <summary>
/// Base Class of Social event
/// </summary>
public abstract class BaseSocialEvent
{
    public enum EventType
    {
        None,
        Test
    }

    // Type of event
    public int _type;

    // The in-game time of this event 
    public uint _time;

    // The in-game place of this event
    public string _place;

    // The sender character of this event
    public Character _sender;

    // The receiver character of this event
    public Character _receiver;

    // Process mothed of this event
    public abstract void Process();

    // Generate a description of this event
    public abstract string Description();
}
