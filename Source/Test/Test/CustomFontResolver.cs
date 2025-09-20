using PdfSharp.Fonts;
using Test.Properties;

namespace Test
{
    public class CustomFontResolver : IFontResolver
    {
        public byte[] GetFont(string faceName)
        {
            if (faceName == "Libre Barcode 39") return Resources.LibreBarcode39_Regular;
            return faceName.EndsWith("#b") ? Resources.NotoSansJP_ExtraBold : Resources.NotoSansJP_Regular;
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            var faceName = familyName; 
            if (isBold) faceName += "#b";
            return new FontResolverInfo(faceName);
        }
    }
}
