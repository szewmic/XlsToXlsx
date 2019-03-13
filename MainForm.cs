using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace XlsToXlsx
{
    public partial class MainForm : Form
    {
        private Transformer transformer;
        private Logger logger;
        private string selectedFolderPath;
        public MainForm()
        {
            logger = new Logger();
            transformer = new Transformer();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();

            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                selectedFolderPath = dialog.FileName;
                label2.Text = "Wybrana ścieżka " + selectedFolderPath;    
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            if (MessageBox.Show("Czy jesteś pewien?", "Potwierdzenie", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(selectedFolderPath))
                {

                    this.progressBar1.Visible = true;

                    var progress = new Progress<int>(v =>
                    {
                        // This lambda is executed in context of UI thread,
                        // so it can safely update form controls
                        progressBar1.Value = v;
                    });

                    Logger.LogMainFolderPath(selectedFolderPath);
                    transformer.MainFolderPath = selectedFolderPath;
                    transformer.MinLength = Convert.ToInt32(numericUpDown1.Value) * 1000000;
                    transformer.DeleteXlsFlag = checkBox1.Checked;

                    // Run operation in another thread
                    await Task.Run(() => transformer.TransformXls_Xlsx(progress));

                    // TODO: Do something after all calculations
                    logger.CreateLogFile();
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
            if (progressBar1.Value < 100 & progressBar1.Value > 0)
            {
                var app = new Microsoft.Office.Interop.Excel.Application();
                
                foreach(Microsoft.Office.Interop.Excel.Workbook w in app.Workbooks)
                {
                    w.Close();
                }

                logger.CreateLogFile();
            }
        }
    }
}
