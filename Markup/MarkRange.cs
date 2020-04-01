using System;
using System.Text.RegularExpressions;


namespace HeaderMarkup.Markup
{
    class MarkRange
    {
        public static readonly Regex AddressRegex = new Regex(@"^\$(?<Left>[A-Z]{1,3})\$(?<Top>\d+)(?::\$(?<Right>[A-Z]{1,3})\$(?<Bottom>\d+))?$", RegexOptions.Compiled);
        public string Address { get; }
        public int left, top, right, bottom;

        public MarkRange(string address)
        {
            Match match = AddressRegex.Match(address);
            if (!match.Success)
                throw new Exception($"Parse Range Fail.\nInvalid Range: {address}.");
            Address = address;
            left = Utils.ParseColumn(match.Groups["Left"].Value);
            top = Convert.ToInt32(match.Groups["Top"].Value);
            if (match.Groups["Right"].Value != string.Empty)
            {
                right = Utils.ParseColumn(match.Groups["Right"].Value);
                bottom = Convert.ToInt32(match.Groups["Bottom"].Value);
            }
            else
            {
                right = left;
                bottom = top;
            }
        }
        public bool IsOverlap(MarkRange range) => !(left > range.right || range.left > right || top > range.bottom || range.top > bottom);
        public bool IsInside(MarkRange range) => (left >= range.left && top >= range.top && right <= range.right && bottom <= range.bottom);
    }
}
