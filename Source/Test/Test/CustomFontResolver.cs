using PdfSharp.Fonts;
using Test.Properties;

namespace Test
{
    public class CustomFontResolver : IFontResolver
    {
        public byte[] GetFont(string faceName)
            //Implement so that you can get as many fonts as you need.
            => faceName.EndsWith("#b") ? Resources.NotoSansJP_ExtraBold : Resources.NotoSansJP_Regular;

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            var faceName = familyName; 
            if (isBold) faceName += "#b";
            return new FontResolverInfo(faceName);
        }
    }
}
