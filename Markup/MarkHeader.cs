

namespace HeaderMarkup.Markup
{
    class MarkHeader : MarkRange
    {
        public int Type { get; }

        public MarkHeader(string address, int type) : base(address) => Type = type;

        public override string Name
            => $"[Hd,{Type},{left},{top},{right},{bottom}]";

        public override string ToString()
            => $"[Hd,{Type},{left},{top},{right},{bottom}]";
    }
}
