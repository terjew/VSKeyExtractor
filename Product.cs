using System;

namespace VSKeyExtractor
{
    internal struct Product
    {
        public string Name { get; }
        public Guid GUID { get; }
        public string MPC { get; }
        public Product(string name, Guid guid, string mpc)
        {
            Name = name;
            GUID = guid;
            MPC = mpc;
        }
    }
}
