using System;

namespace DAProcessor
{
    public class Message
    {
        public string message { get; set; }
        
        private DateTime _timestamp;

        public DateTime timestamp
        {
            get { return _timestamp; }
            private set { _timestamp = DateTime.Now; }
        }
        
        public Message()
        {

        }
    }
}
