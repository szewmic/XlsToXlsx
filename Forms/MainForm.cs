using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using XlsToXlsx.Domain;
using XlsToXlsx.Interfaces;

namespace XlsToXlsx
{
    public partial class MainForm : Form
    {
        private readonly Transformer _transformer;
        private readonly SimpleTransformerFactory _transformerFactory;
        private readonly Logger _logger;

        public MainForm(SimpleTransformerFactory transformerFactory)
        {
            _transformerFactory = transformerFactory;
            _transformer = _transformerFactory.CreateTransformer("excel");
            _logger = Logger.getInstance();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();

            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                _transformer.MainFolderPath = dialog.FileName;
                label2.Text = "Wybrana ścieżka " + _transformer.MainFolderPath;    
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            if (MessageBox.Show("Czy jesteś pewien?", "Potwierdzenie", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(_transformer.MainFolderPath))
                {

                    this.progressBar1.Visible = true;

                    var progress = new Progress<int>(v =>
                    {
                        // This lambda is executed in context of UI thread,
                        // so it can safely update form controls
                        progressBar1.Value = v;
                    });

                    _logger.LogMainFolderPath(_transformer.MainFolderPath);
                    _transformer.MinLength = Convert.ToInt32(numericUpDown1.Value) * 1000000;
                    _transformer.DeleteFileFlag = checkBox1.Checked;

                    // Run operation in another thread
                    await Task.Run(() => _transformer.Transform(progress));

                    // TODO: Do something after all calculations
                    _logger.CreateLogFile();
                    MessageBox.Show("Operacja zakończona pomyślnie!\n Utworzono log na pulpicie!");

                    this.progressBar1.Visible = false;
                    this.progressBar1.Value = 0;
                }
                else
                    MessageBox.Show("Musisz wskazać scieżkę!");
            }
            button2.Enabled = true;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (progressBar1.Value > 0 && progressBar1.Value < 100)
            {
                _transformer.Interrupt();
            }
        }
    }
}
