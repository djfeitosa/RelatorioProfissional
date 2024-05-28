using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GeradorRelatorioPDF
{
    [Serializable]
    public class Profissao
    {
        public int IdProfissao { get; set; }
        public string Nome { get; set; }

    }
}