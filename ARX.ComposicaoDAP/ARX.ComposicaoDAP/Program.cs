using ARX.ComposicaoDAP.Domain;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ARX.ComposicaoDAP
{
    class Program
    {
        private static readonly DateTime dataInicioFundo = new DateTime(2018, 1, 2);
        private static readonly DateTime dataFimFundo = new DateTime(2020, 1, 17);
        private static readonly decimal valorFundo = 10000000m;
        private static readonly int qtdLotePadrao = 5;

        static void Main(string[] args)
        {
            var contratos = CarregarContratosFromExcel();

            EncontrarComposicaoCarteira(contratos);
        }

        private static ICollection<ContratoDAP> CarregarContratosFromExcel()
        {
            Console.WriteLine("Carregando arquivo....");

            using var file = new XLWorkbook(Environment.CurrentDirectory + @"\historico_dap.xlsx");
            var sheet = file.Worksheets.First();
            var totalLinhas = sheet.Rows().Count();
            var contratos = new List<ContratoDAP>(totalLinhas);

            Console.WriteLine($"Arquivo carregado, {totalLinhas} linhas");
            Console.WriteLine($"Carregando contratos...");

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

            Console.WriteLine($"Contratos carregados");

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

            return codigos;
        }

        private static ICollection<RentabilidadeFundoDAP> ConstruirRentabilidades(
            IEnumerable<string> codigos, ICollection<ContratoDAP> contratos)
        {
            Console.WriteLine($"Identificando rentabilidades...\r\n");

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
            decimal valorLastro = valorFundo;
            var carteira = new Dictionary<ContratoDAP, int>();

            foreach (var item in rentabilidades.OrderByDescending(x => x.Rentabilidade)
                .Where(x => x.ContratoInicio.DataReferencia.Equals(dataInicioFundo)
                    && x.ContratoFim.DataReferencia.Equals(dataFimFundo)))
            {
                decimal valorLote = item.ContratoInicio.PU * qtdLotePadrao;
                int qtdContratos = 0;

                while (valorLote <= valorLastro)
                {
                    qtdContratos += qtdLotePadrao;
                    valorLastro -= valorLote;
                }

                carteira.Add(item.ContratoInicio, qtdContratos);

                Console.WriteLine($"DAP: {item.ContratoInicio.Codigo} | PU Inic: {item.ContratoInicio.PU} | PU Fim: {item.ContratoFim.PU} " +
                    $"| Rentabilidade: {item.Rentabilidade} | Dt. Venc: {item.ContratoInicio.DataVencimento} " +
                    $"| Valor lote: {valorLote}");
            }

            Console.WriteLine($"\r\nValor Fundo: {valorFundo}");
            Console.WriteLine($"Valor restando: {valorLastro}");

            return carteira;
        }

        private static void ImprimirCarteira(Dictionary<ContratoDAP, int> carteira)
        {
            Console.WriteLine("\r\nDistribuição Carteira!!!\r\n");

            foreach (var item in carteira.Where(x => x.Value > 0))
            {
                Console.WriteLine($"DAP: {item.Key.Codigo} | PU Entrada: {item.Key.PU} | Data Ref: {item.Key.DataReferencia} " +
                    $"| Qtd. Contratos: {item.Value} | Valor Alocado: {item.Key.PU * item.Value}");
            }

            Console.WriteLine($"\r\nExposição Total: {carteira.Sum(x => (x.Key.PU * x.Value) / valorFundo)}");

        }
    }
}
