using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using XlsToXlsx.Interfaces;

namespace XlsToXlsx
{
    class ExcelTransformer : Transformer
    {
        public ExcelTransformer() { }

        public override void Transform(IProgress<int> progress)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var FilesPaths = LoadFilesPaths("xls");
            var progressStep = 100/FilesPaths.Count;
            var progressTotal = 0;

            foreach (string oldXlsPath in FilesPaths)
            {
                Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open(oldXlsPath);
                var newXlsxPath = Path.ChangeExtension(oldXlsPath, ".xlsx");

                SaveXlsAsXlsx(workbook, newXlsxPath);
                workbook.Close();
                CountCreatedWeight(newXlsxPath);

                if (DeleteFileFlag)
                {
                    CountDeletedWeight(oldXlsPath);
                    DeleteOldFile(oldXlsPath, newXlsxPath);
                }


                if (progress != null)
                    progress.Report(progressTotal);
                progressTotal = progressTotal + progressStep;
            }
            Logger.getInstance().LogXlsDeletedAndXlsxCreatedWeights(deletedWeight, createdWeight);

            createdWeight = 0;
            deletedWeight = 0;
            FilesPaths.Clear();
            app.Quit();
        }

        private void SaveXlsAsXlsx(Microsoft.Office.Interop.Excel.Workbook passedWorkbook, string newXlsxPath)
        {
            try
            {
                passedWorkbook.SaveAs(Filename: newXlsxPath, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                Logger.getInstance().LogNewFilePath(newXlsxPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

    }
}
