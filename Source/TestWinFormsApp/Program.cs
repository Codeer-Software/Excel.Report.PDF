using PdfSharp.Fonts;
using Test;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace TestWinFormsApp
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            GlobalFontSettings.FontResolver = new CustomFontResolver();

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            var main = new MainForm();
            main.StartPosition = FormStartPosition.CenterScreen;
            Application.Run(main);
        }
    }
}