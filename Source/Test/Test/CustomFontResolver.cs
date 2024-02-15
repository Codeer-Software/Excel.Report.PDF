using PdfSharp.Fonts;
using Test.Properties;

namespace Test
{
    public class CustomFontResolver : IFontResolver
    {
        public byte[] GetFont(string faceName)
            //Implement so that you can get as many fonts as you need.
            => Resources.NotoSansJP_VariableFont_wght;

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
            => new FontResolverInfo(familyName);
    }
}
