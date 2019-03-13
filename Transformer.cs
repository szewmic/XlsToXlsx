using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace XlsToXlsx
{
    class Transformer
    {
        private string _mainFolderPath;
        private int _minLength;

        private long xlsDeletedWeight = 0;
        private long XlsxCreatedWeight = 0;
        public bool DeleteXlsFlag { get; set; }

        public string MainFolderPath
        {
            get
            {
                return _mainFolderPath;
            }
            set
            {
                if(!string.IsNullOrEmpty(value))
                _mainFolderPath = value;
            }
        }

        public int MinLength
        {
            get
            {
                return _minLength;
            }
            set
            {
                if (value > 0)
                    _minLength = value;
            }
        }

        public Transformer() { }
        public Transformer(string mainFolderPath, int minLength)
        {
            this.MainFolderPath = mainFolderPath;
            this.MinLength = minLength;
        }

        private List<string> LoadFilesPaths()
        {
            DirectoryInfo dir = new DirectoryInfo(_mainFolderPath);

            IEnumerable<FileInfo> files = dir.GetFiles("*.xls", SearchOption.AllDirectories);

            return files.Where(s => s.Length >= _minLength && s.Extension == ".xls").Select(s => s.FullName).ToList();
        }

        public void TransformXls_Xlsx(IProgress<int> progress)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var FilesPaths = LoadFilesPaths();
            var progressStep = 100/FilesPaths.Count;
            var progressTotal = 0;

            foreach (string oldXlsPath in FilesPaths)
            {
                Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open(oldXlsPath);
                var newXlsxPath = Path.ChangeExtension(oldXlsPath, ".xlsx");

                SaveXlsAsXlsx(workbook, newXlsxPath);
                workbook.Close();
                CountXlsxCreatedWeight(newXlsxPath);

                if (DeleteXlsFlag)
                {
                    CountXlsDeletedWeight(oldXlsPath);
                    DeleteOldXlsFile(oldXlsPath, newXlsxPath);
                }


                if (progress != null)
                    progress.Report(progressTotal);
                progressTotal = progressTotal + progressStep;


            }

            Logger.LogXlsDeletedAndXlsxCreatedWeights(xlsDeletedWeight, XlsxCreatedWeight);

            XlsxCreatedWeight = 0;
            xlsDeletedWeight = 0;
            FilesPaths.Clear();
            app.Quit();
        }

        private void SaveXlsAsXlsx(Microsoft.Office.Interop.Excel.Workbook passedWorkbook, string newXlsxPath)
        {
            try
            {
                passedWorkbook.SaveAs(Filename: newXlsxPath, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                Logger.LogNewFilePath(newXlsxPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
        private void DeleteOldXlsFile(string oldXlsPath, string newXlsxPath)
        {
            if (File.Exists(oldXlsPath) && File.Exists(newXlsxPath))
            {
                try
                {
                    File.Delete(oldXlsPath);
                    Logger.LogDeletedFilePath(oldXlsPath);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }
        }

        private void CountXlsDeletedWeight(string path) => xlsDeletedWeight = xlsDeletedWeight + new FileInfo(path).Length;
        private void CountXlsxCreatedWeight(string path) => XlsxCreatedWeight = XlsxCreatedWeight + new FileInfo(path).Length;
    }
}
