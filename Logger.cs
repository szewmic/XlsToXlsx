using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XlsToXlsx
{
    class Logger
    {
        private static string mainFolderPath;
        private static List<string> deletedFilePaths = new List<string>();
        private static List<string> newFilePaths = new List<string>();
        private static long xlsDeletedWeight = 0;
        private static long xlsxCreatedWeight = 0;

        public static void LogMainFolderPath(string path)
        {
            if (!string.IsNullOrEmpty(path))
                mainFolderPath = path;
        }
        public static void LogDeletedFilePath(string path)
        {
            if (!string.IsNullOrEmpty(path))
                deletedFilePaths.Add(path);
        }
        public static void LogNewFilePath(string path)
        {
            if (!string.IsNullOrEmpty(path))
                newFilePaths.Add(path);
        }
        public static void LogXlsDeletedAndXlsxCreatedWeights(long xlsDel, long xlsxCrea) { xlsDeletedWeight = xlsDel; xlsxCreatedWeight = xlsxCrea; }

        public void CreateLogFile()
        {
            try
            {
                File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    + $"\\XLS log {DateTime.Now.ToString("yyyyMMdd_HHmmss")}.txt"
                    , GetLog());
            } catch (UnauthorizedAccessException ee)
            {
                MessageBox.Show(ee.ToString());
            } catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            };
            ClearLogger();
        }
        private void ClearLogger()
        {
            mainFolderPath = "";
            deletedFilePaths.Clear();
            newFilePaths.Clear();
        }
        private string GetLog()
        {
             return MakeHeader() +"\n"
                + MakeNewFileSection()+"\n"
                + MakeDeletedFileSection();
        }
        private string MakeHeader()
        {
            string header = $"Log utworzony przez program 'XlsToXlsx' \n" 
                + $"Log utworzony dnia: {DateTime.Now}\n"
                + $"Nadrzędny wybrany folder to: {mainFolderPath}\n"
                +$"Waga utworzonych plików .XLSX: {xlsxCreatedWeight/1000000} [MB]\n"
                + $"Waga usuniętych plików .XLS: {xlsDeletedWeight/1000000} [MB]\n";

            if (xlsDeletedWeight > 0)
                header = header +
                    $"Odzyskane miejsce na dysku {Math.Abs(xlsxCreatedWeight - xlsDeletedWeight)/1000000} [MB]\n";

            return header;
        }
        private string MakeDeletedFileSection()
        {
            string section = "Poniżej znajduję się lista wszystkich usuniętych plików .xls: \n";
            foreach (string path in deletedFilePaths)
            {
                section = section + path + "\n";
            }
            return section;
        }
        private string MakeNewFileSection()
        {
            string section = "Poniżej znajduję się lista wszystkich utworzonych plików .xlsx: \n";
            foreach (string path in newFilePaths)
            {
                section = section + path + "\n";
            }
            return section;
        }

    }
}
