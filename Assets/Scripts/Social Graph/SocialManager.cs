using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class SocialManager : MonoBehaviour
{
    protected static SocialManager _instance;

    public static SocialManager Instance
    {
        get
        {
            return _instance;
        }
    }

    // Factor to control the amount of events that can be handled in each frame
    [Range(0, 100)]
    public int _eventPerFrame = 5;

    // Hold events happaned in each frame
    private Queue<BaseSocialEvent> _eventQueue;

    #region Propertise

    public SocialGraph Graph { get; private set; }

    #endregion

    // Start is called before the first frame update
    void Awake()
    {
        if (_instance)
        {
            Debug.LogError("Only one instance of Social Manager can exist.");
        }
        else
        {
            Graph = new SocialGraph();
            _eventQueue = new Queue<BaseSocialEvent>();
        }
    }

    // Update is called once per frame
    void Update()
    {
        // Handle event
        if (_eventQueue.Count > 0)
        {
            for (int i = 0; i < _eventPerFrame; ++i)
            {
                var socialEvent = _eventQueue.Peek();
                socialEvent.Process();
                _eventQueue.Dequeue();
            }
        }
    }

    public void RegisterSocialEvent(BaseSocialEvent socialEvent)
    {
        _eventQueue.Enqueue(socialEvent);
    }
}
