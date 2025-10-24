using Microsoft.Win32;
using PdfSharp.Fonts;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Versioning;

namespace TestWinFormsApp
{
    [SupportedOSPlatform("windows")]
    internal sealed class WindowsInstalledFontResolver : IFontResolver
    {
        // Example: If Japanese is primary, search in an order such as Yu Gothic UI, then Segoe UI, etc.
        private readonly string[] _fallbackFamilies;

        // Value name → absolute path to the font file (.ttf/.otf only)
        private readonly Dictionary<string, string> _fontNameToPath;

        // faceName → font bytes
        private readonly ConcurrentDictionary<string, byte[]> _cache = new(StringComparer.OrdinalIgnoreCase);

        private static readonly string WindowsFontsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

        public WindowsInstalledFontResolver(params string[]? fallbackFamilies)
        {
            _fallbackFamilies = (fallbackFamilies is { Length: > 0 })
                ? fallbackFamilies
                : new[] { "Segoe UI", "Yu Gothic UI", "Meiryo UI", "MS UI Gothic" };

            _fontNameToPath = LoadFontRegistryMap();
        }

        /// <summary>
        /// Called each time PdfSharp requests the actual font bytes.
        /// faceName is the value returned by ResolveTypeface.
        /// </summary>
        public byte[] GetFont(string faceName)
        {
            // Cache
            if (_cache.TryGetValue(faceName, out var cached))
                return cached;

            // Assumes faceName is encoded as "familyName[#b][#i]"
            ParseFaceName(faceName, out var family, out var wantBold, out var wantItalic);

            if (!TryFindFontPath(family, wantBold, wantItalic, out var path))
            {
                // Fallback
                foreach (var fb in _fallbackFamilies)
                {
                    if (TryFindFontPath(fb, wantBold, wantItalic, out path))
                        break;
                }
            }

            if (path is null)
                throw new FileNotFoundException($"Installed font not found for '{faceName}' (family='{family}', bold={wantBold}, italic={wantItalic}).");

            var bytes = File.ReadAllBytes(path);
            _cache[faceName] = bytes;
            return bytes;
        }

        /// <summary>
        /// Determine faceName from familyName and style.
        /// You can return any string to PdfSharp, but it must match what GetFont expects.
        /// </summary>
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // PdfSharp doesn't care much about the case of familyName, but we preserve it here.
            // The internal representation of faceName is "family#b#i".
            var face = BuildFaceName(familyName, isBold, isItalic);
            return new FontResolverInfo(face);
        }

        private static string BuildFaceName(string family, bool bold, bool italic)
        {
            var f = family;
            if (bold) f += "#b";
            if (italic) f += "#i";
            return f;
        }

        private static void ParseFaceName(string faceName, out string family, out bool bold, out bool italic)
        {
            bold = faceName.Contains("#b", StringComparison.OrdinalIgnoreCase);
            italic = faceName.Contains("#i", StringComparison.OrdinalIgnoreCase);
            family = faceName
                .Replace("#b", "", StringComparison.OrdinalIgnoreCase)
                .Replace("#i", "", StringComparison.OrdinalIgnoreCase)
                .Trim();
        }

        /// <summary>
        /// Build a dictionary from the Windows font registry for .ttf/.otf only.
        /// Value name (e.g., "Segoe UI Bold (TrueType)") → absolute path.
        /// </summary>
        private static Dictionary<string, string> LoadFontRegistryMap()
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts");
            if (key is null) return map;

            foreach (var valueName in key.GetValueNames())
            {
                var v = key.GetValue(valueName) as string;
                if (string.IsNullOrWhiteSpace(v))
                    continue;

                // The path may be relative (under the Fonts directory) or absolute.
                var path = v;
                if (!Path.IsPathRooted(path))
                    path = Path.Combine(WindowsFontsDir, path);

                // Skip TTC (often problematic with PdfSharp).
                var ext = Path.GetExtension(path).ToLowerInvariant();
                if (ext is ".ttf" or ".otf")
                {
                    // If duplicates exist with the same name, prefer the first one.
                    if (!map.ContainsKey(valueName))
                        map[valueName] = path;
                }
            }
            return map;
        }

        /// <summary>
        /// Find the most suitable registry entry from family + style and return the path.
        /// </summary>
        private bool TryFindFontPath(string family, bool bold, bool italic, [NotNullWhen(true)] out string? path)
        {
            path = null;

            // The registry value names tend to look like the following:
            // "Segoe UI (TrueType)"
            // "Segoe UI Bold (TrueType)"
            // "Segoe UI Italic (TrueType)"
            // "Segoe UI Bold Italic (TrueType)"
            // etc. The part in parentheses like "(TrueType)" is not always consistent, so we search by partial match.

            static IEnumerable<string> Candidates(string fam, bool b, bool i)
            {
                var baseKey = fam + " ";
                if (b && i)
                    yield return fam + " Bold Italic";
                if (b && !i)
                    yield return fam + " Bold";
                if (!b && i)
                    yield return fam + " Italic";
                yield return fam; // Regular
                // For UI families, sometimes only "Semibold" exists, so we add this conservatively.
                if (b) yield return fam + " SemiBold";
            }

            foreach (var cand in Candidates(family, bold, italic))
            {
                // Loosely match by prefix of the value name (e.g., starts with "Segoe UI Bold").
                var hit = _fontNameToPath.Keys.FirstOrDefault(k => k.StartsWith(cand, StringComparison.OrdinalIgnoreCase));
                if (hit is not null)
                {
                    var p = _fontNameToPath[hit];
                    if (File.Exists(p))
                    {
                        path = p;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
