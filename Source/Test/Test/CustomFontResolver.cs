using PdfSharp.Fonts;
using Test.Properties;

namespace Test
{
    public class CustomFontResolver : IFontResolver
    {
        public byte[] GetFont(string faceName)
            => Resources.ipag;

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
            => new FontResolverInfo(familyName);
    }
}
