namespace Test
{
    public static class TestEnvironment
    {
        public static string PdfSrcPath => Path.Combine(GetDirectory(typeof(TestEnvironment).Assembly.Location, 5), "Data");
        public static string TestResultsPath => Path.Combine(GetDirectory(typeof(TestEnvironment).Assembly.Location, 5), "Results");

        static string GetDirectory(string path, int count)
        {
            for (int i = 0; i < count; i++) path = Path.GetDirectoryName(path)!;
            return path;
        }
    }
}