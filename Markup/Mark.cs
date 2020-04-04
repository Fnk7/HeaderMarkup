

namespace HeaderMarkup.Markup
{
    class Mark : Range
    {
        public int Type { get; }

        public Mark(string address, int type) : base(address) => Type = type;

        public override string Name
            => $"[Mk,{Type},{left},{top},{right},{bottom}]";

        public override string ToString()
            => $"[Mk,{Type},{left},{top},{right},{bottom}]";
    }
}
