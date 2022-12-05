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
            var contratos = CarregarExcel();

            EncontrarComposicaoCarteira(contratos);
        }

        private static ICollection<ContratoDAP> CarregarExcel()
        {
            Console.WriteLine("Carregando arquivo....");

            var file = new XLWorkbook(@"C:\Users\Marco\source\repos\ARX.ComposicaoDAP\ARX.ComposicaoDAP\historico_dap.xlsx");
            var sheet = file.Worksheets.First();
            var totalLinhas = sheet.Rows().Count();
            var contratos = new List<ContratoDAP>(totalLinhas);

            Console.WriteLine($"Arquivo carregado, {totalLinhas} linhas");
            Console.WriteLine($"Carregando dados");

            for (int linha = 2; linha <= totalLinhas; linha++)
            {
                contratos.Add(new ContratoDAP
                {
                    Codigo = sheet.Cell($"C{linha}").Value.ToString(),
                    PU = decimal.Parse(sheet.Cell($"E{linha}").Value.ToString()),
                    DataReferencia = DateTime.Parse(sheet.Cell($"A{linha}").Value.ToString())
                });

            }

            Console.WriteLine($"Dados carregado");

            return contratos;
        }

        private static void EncontrarComposicaoCarteira(ICollection<ContratoDAP> contratos)
        {
            var codigos = IdentificarCodigos(contratos);

            var rentabilidades = IdentificarRentabilidades(codigos, contratos);

            var carteira = ComporCarteira(rentabilidades);
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

        private static Dictionary<ContratoDAP, decimal> IdentificarRentabilidades(
            IEnumerable<string> codigos, ICollection<ContratoDAP> contratos)
        {
            Console.WriteLine($"Identificando rentabilidades...");

            Dictionary<ContratoDAP, decimal> rentabilidades =
                new Dictionary<ContratoDAP, decimal>();

            foreach (var codigo in codigos)
            {
                var referenciaInicio = contratos.Where(x => x.Codigo.Equals(codigo)
                                            && x.DataReferencia.Equals(new DateTime(2018, 1, 2)))
                    .FirstOrDefault();

                var referenciaFim = contratos.Where(x => x.Codigo.Equals(codigo)
                                            && x.DataReferencia.Equals(new DateTime(2020, 1, 17)))
                    .FirstOrDefault();


                if (referenciaInicio != null && referenciaFim != null)
                {
                    var rentabilidade = ((referenciaFim.PU - referenciaInicio.PU) * 100) / referenciaInicio.PU;

                    rentabilidades.Add(referenciaFim, rentabilidade);
                }

            }

            return rentabilidades;
        }

        private static Dictionary<ContratoDAP, int> ComporCarteira(Dictionary<ContratoDAP, decimal> rentabilidades)
        {
            decimal valorLastro = 10100000m;
            int qtdLote = 5;
            var carteira = new Dictionary<ContratoDAP, int>();

            foreach (var item in rentabilidades.OrderByDescending(x => x.Value))
            {
                decimal valorLote = item.Key.PU * qtdLote;
                int qtdContratos = 0;

                while (valorLote <= valorLastro)
                {
                    qtdContratos += qtdLote;
                    valorLastro -= valorLote;
                }

                carteira.Add(item.Key, qtdContratos);

                Console.WriteLine($"DAP: {item.Key.Codigo} | Rentabilidade: {item.Value} " +
                    $"| PU: {item.Key.PU} | Valor lote: {valorLote} | Contratos: {qtdContratos}");
            }

            Console.WriteLine($"Valor restando: {valorLastro}");
            
            return carteira;
        }
    }
}
