using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XlsToXlsx.Interfaces
{
     abstract class Transformer
    {
        protected string _mainFolderPath;
        protected int _minLength;

        protected long deletedWeight = 0;
        protected long createdWeight = 0;

        public bool DeleteFileFlag { get; set; }

        public string MainFolderPath
        {
            get
            {
                return _mainFolderPath;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
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

        public abstract void Transform(IProgress<int> progress);
        protected List<string> LoadFilesPaths(string extensionToLoad)
        {
            DirectoryInfo dir = new DirectoryInfo(_mainFolderPath);

            IEnumerable<FileInfo> files = dir.GetFiles($"*.{extensionToLoad}", SearchOption.AllDirectories);

            return files.Where(s => s.Length >= _minLength && s.Extension == $".{extensionToLoad}").Select(s => s.FullName).ToList();
        }
        protected void DeleteOldFile(string oldPath, string newPath)
        {
            if (File.Exists(oldPath) && File.Exists(newPath))
            {
                try
                {
                    File.Delete(oldPath);
                    Logger.getInstance().LogDeletedFilePath(oldPath);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }
        }
        protected void CountDeletedWeight(string path) => deletedWeight = deletedWeight + new FileInfo(path).Length;
        protected void CountCreatedWeight(string path) => createdWeight = createdWeight + new FileInfo(path).Length;
    }
}
