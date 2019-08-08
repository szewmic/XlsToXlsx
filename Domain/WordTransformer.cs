using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XlsToXlsx.Interfaces;


namespace XlsToXlsx.Domain
{
    public class WordTransformer : Transformer
    {
        public override void Interrupt()
        {
            throw new NotImplementedException();
        }

        public override void Transform(IProgress<int> progress)
        {
            throw new NotImplementedException();
        }
    }
}
