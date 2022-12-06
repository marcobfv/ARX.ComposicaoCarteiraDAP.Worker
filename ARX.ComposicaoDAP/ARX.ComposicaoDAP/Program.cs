using ARX.ComposicaoDAP.Domain;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ARX.ComposicaoDAP
{
    class Program
    {
        static void Main(string[] args)
        {
            var contratos = CarregarContratosFromExcel();

            EncontrarComposicaoCarteira(contratos);
        }

        private static ICollection<ContratoDAP> CarregarContratosFromExcel()
        {
            Console.WriteLine("Carregando arquivo....");

            using var file = new XLWorkbook(@"C:\Users\Marco\source\repos\ARX.ComposicaoDAP\ARX.ComposicaoDAP\historico_dap.xlsx");
            var sheet = file.Worksheets.First();
            var totalLinhas = sheet.Rows().Count();
            var contratos = new List<ContratoDAP>(totalLinhas);
            var dataInicioFundo = new DateTime(2018, 1, 2);
            var dataFimFundo = new DateTime(2020, 1, 17);

            Console.WriteLine($"Arquivo carregado, {totalLinhas} linhas");
            Console.WriteLine($"Carregando dados");

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                var dataReferencia = DateTime.Parse(sheet.Cell($"A{linha}").Value.ToString());

                if (dataReferencia >= dataInicioFundo
                    && dataReferencia <= dataFimFundo)
                    contratos.Add(new ContratoDAP
                    {
                        Codigo = sheet.Cell($"C{linha}").Value.ToString(),
                        PU = decimal.Parse(sheet.Cell($"E{linha}").Value.ToString()),
                        DataVencimento = DateTime.Parse(sheet.Cell($"D{linha}").Value.ToString()),
                        DataReferencia = dataReferencia
                    });
            }

            Console.WriteLine($"Dados carregado");

            return contratos.OrderBy(x => x.DataReferencia)
                .ToList();
        }

        private static void EncontrarComposicaoCarteira(ICollection<ContratoDAP> contratos)
        {
            var codigos = IdentificarCodigos(contratos);

            var rentabilidades = ConstruirRentabilidades(codigos, contratos);

            var carteira = ComporCarteira(rentabilidades);

            ImprimirCarteira(carteira);
        }

        private static IEnumerable<string> IdentificarCodigos(ICollection<ContratoDAP> contratos)
        {
            Console.WriteLine($"Identificando codigos...");

            var codigos = contratos.GroupBy(x => x.Codigo)
                            .Select(grp => grp.Key)
                            .ToList();

            foreach (var item in codigos)
            {
                Console.WriteLine($"Codigo: {item}");
            }

            return codigos;
        }

        private static ICollection<RentabilidadeFundoDAP> ConstruirRentabilidades(
            IEnumerable<string> codigos, ICollection<ContratoDAP> contratos)
        {
            Console.WriteLine($"Identificando rentabilidades...");

            ICollection<RentabilidadeFundoDAP> rentabilidades =
                new List<RentabilidadeFundoDAP>();

            foreach (var codigo in codigos)
            {
                var referenciaInicio = contratos.Where(x => x.Codigo.Equals(codigo)).First();
                var referenciaFim = contratos.Where(x => x.Codigo.Equals(codigo)).Last();

                rentabilidades.Add(new RentabilidadeFundoDAP(referenciaInicio, referenciaFim));
            }

            return rentabilidades;
        }

        private static Dictionary<ContratoDAP, int> ComporCarteira(ICollection<RentabilidadeFundoDAP> rentabilidades)
        {
            decimal valorFundo = 10000000m;
            decimal valorLastro = valorFundo;
            int qtdLote = 20;
            var carteira = new Dictionary<ContratoDAP, int>();

            foreach (var item in rentabilidades.OrderByDescending(x => x.Rentabilidade)
                .Where(x => x.ContratoInicio.DataReferencia.Equals(new DateTime(2018, 1, 2))))
            {
                decimal valorLote = item.ContratoInicio.PU * qtdLote;
                int qtdContratos = 0;

                while (valorLote <= valorLastro)
                {
                    qtdContratos += qtdLote;
                    valorLastro -= valorLote;
                }

                carteira.Add(item.ContratoInicio, qtdContratos);

                Console.WriteLine($"DAP: {item.ContratoInicio.Codigo} | PU Inic: {item.ContratoInicio.PU} | PU Fim: {item.ContratoFim.PU} " +
                    $"| Rentabilidade: {item.Rentabilidade} | Dt. Venc: {item.ContratoInicio.DataVencimento} " +
                    $"| Valor lote: {valorLote}");
            }

            Console.WriteLine($"Valor Fundo: {valorFundo}");
            Console.WriteLine($"Valor restando: {valorLastro}");
            Console.WriteLine($"Exposição Total: {carteira.Sum(x => (x.Key.PU * x.Value) / valorFundo)}");

            return carteira;
        }

        private static void ImprimirCarteira(Dictionary<ContratoDAP, int> carteira)
        {
            Console.WriteLine("\r\nDistribuição Carteira!!!\r\n");

            foreach (var item in carteira.Where(x => x.Value > 0))
            {
                Console.WriteLine($"DAP: {item.Key.Codigo} | Qtd. Contratos: {item.Value}");
            }
        }
    }
}
