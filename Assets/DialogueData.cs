using System.Collections;
using System.Collections.Generic;
using UnityEngine;
namespace Dialogue
{
    [CreateAssetMenu(fileName = "DialogueSo")]
    public class DialogueData : ScriptableObject
    {
        public List<Sentence> data = new List<Sentence>();

        [System.Serializable]
        public class Sentence
        {
            public string character;
            [TextArea]
            public string content;
        }
    }
}