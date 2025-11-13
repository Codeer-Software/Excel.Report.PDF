using Excel.Report.PrintDocument;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace TestWinFormsApp
{
    public partial class MainForm : Form
    {
        private readonly PrintDocument _doc = new PrintDocument();
        private readonly PrintPreviewDialog _preview = new PrintPreviewDialog();
        private readonly PageSetupDialog _pageSetup = new PageSetupDialog();

        public MainForm()
        {
            InitializeComponent();
            _preview.Document = _doc;
            _pageSetup.Document = _doc;
            _preview.Width = 1000;
            _preview.Height = 800;
        }

        private void _buttonFile_Click(object sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog();
            ofd.Filter = "Excelファイル|*.xlsx;*.xlsm;*.xlsb;*.xls|すべてのファイル|*.*";
            if (ofd.ShowDialog() != DialogResult.OK) return;
            _textBoxFile.Text = ofd.FileName;
        }

        private void _buttonPreview_Click(object sender, EventArgs e)
        {
            ExcelPrintDocumentBinder.Bind(_doc, _textBoxFile.Text);
            _preview.StartPosition = FormStartPosition.CenterScreen;
            _preview.ShowDialog(this);
        }

        private void _buttonSettings_Click(object sender, EventArgs e)
        {
            _pageSetup.ShowDialog(this);
        }
    }
}
