using System;
using System.Collections.Generic;
using System.Text;

namespace ARX.ComposicaoDAP.Domain
{
    public class ContratoDAP
    {
        public string Codigo { get; set; }

        public decimal PU { get; set; }

        public DateTime DataReferencia { get; set; }

        public DateTime DataVencimento { get; set; }

    }
}
