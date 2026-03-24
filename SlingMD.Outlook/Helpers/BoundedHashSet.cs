using System;
using System.Collections.Generic;

namespace SlingMD.Outlook.Helpers
{
    /// <summary>
    /// Thread-safe bounded deduplication set with LRU eviction.
    /// Evicts the oldest entry when capacity is reached.
    /// </summary>
    public class BoundedHashSet
    {
        private readonly int _capacity;
        private readonly HashSet<string> _set;
        private readonly LinkedList<string> _order;
        private readonly object _lock = new object();

        /// <summary>
        /// Initializes a new BoundedHashSet.
        /// </summary>
        /// <param name="capacity">Maximum number of entries before eviction begins. Defaults to 10000.</param>
        /// <param name="comparer">String comparer to use. Defaults to StringComparer.OrdinalIgnoreCase.</param>
        public BoundedHashSet(int capacity = 10000, IEqualityComparer<string> comparer = null)
        {
            _capacity = capacity;
            IEqualityComparer<string> resolvedComparer = comparer ?? StringComparer.OrdinalIgnoreCase;
            _set = new HashSet<string>(resolvedComparer);
            _order = new LinkedList<string>();
        }

        /// <summary>
        /// Returns true if the item is in the set.
        /// </summary>
        public bool Contains(string item)
        {
            lock (_lock)
            {
                return _set.Contains(item);
            }
        }

        /// <summary>
        /// Adds the item to the set. If already present, returns false.
        /// If at capacity, evicts the oldest entry before adding.
        /// Returns true if the item was newly added.
        /// </summary>
        public bool Add(string item)
        {
            lock (_lock)
            {
                if (_set.Contains(item))
                {
                    return false;
                }

                if (_set.Count >= _capacity)
                {
                    string oldest = _order.First.Value;
                    _order.RemoveFirst();
                    _set.Remove(oldest);
                }

                _set.Add(item);
                _order.AddLast(item);
                return true;
            }
        }

        /// <summary>
        /// Gets the current number of items in the set.
        /// </summary>
        public int Count
        {
            get
            {
                lock (_lock)
                {
                    return _set.Count;
                }
            }
        }

        /// <summary>
        /// Removes all items from the set.
        /// </summary>
        public void Clear()
        {
            lock (_lock)
            {
                _set.Clear();
                _order.Clear();
            }
        }
    }
}
