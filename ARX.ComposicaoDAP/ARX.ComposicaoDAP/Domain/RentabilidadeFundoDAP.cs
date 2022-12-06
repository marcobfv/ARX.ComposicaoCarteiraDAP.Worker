using System;
using System.Collections.Generic;
using System.Text;

namespace ARX.ComposicaoDAP.Domain
{
    public class RentabilidadeFundoDAP
    {
        public ContratoDAP ContratoInicio { get; set; }

        public ContratoDAP ContratoFim { get; set; }

        public decimal Rentabilidade { get
            {
                return ((ContratoFim.PU - ContratoInicio.PU) * 100) / ContratoInicio.PU;
            }
        }

        public RentabilidadeFundoDAP(ContratoDAP contratoInicio, ContratoDAP contratoFim)
        {
            ContratoInicio = contratoInicio;
            ContratoFim = contratoFim;
        }
    }
}
