using ConciliacaoEstoque;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTesteConciliacao
{
    class Program
    {
        static void Main(string[] args)
        {

            // string pathPlanilhaEstoque = @"C:\planilhasEstoque\01_CONTA_CCL.xlsx";
            //string pathPlanilhaRazao = @"C:\planilhasEstoque\01_CONTA_CCL.xlsx";

            var pathPlanilhaFiliais = @"C:\planilhasEstoque\rz_todas_filias.xlsx";
            var pathPlanilha = @"C:\planilhasEstoque\01_CONTA_CCL_CMPLETA.xlsx";
            var pathPlanilhaRazao = @"C:\planilhasEstoque\rz_todas_filias_1.xlsx";

            string palavraChave = "SD2-TOTAL";
            string rangeProcuraPalavra = "a:a";

            // deleta linhas que não serão utilizadas na conciliação
            //Conciliacao.DeletaLinhasPlanilha(pathPlanilha, palavraChave, rangeProcuraPalavra);


            int colunaProcura = 11;
            string rangeProcura = "s:s";
            int linhaInicio = 2;

            //delata linha de notas numero e custo iguais
            //Conciliacao.LinhasValorDuplicado(pathPlanilha, rangeProcura, linhaInicio);

     
            //Conciliacao.ExtraiFilial(pathPlanilhaFiliais, linhaInicio);


            int numColunaOrigem = 10;
            int numColunaDestino = 8;
            //int linhaInicio = 2;

            // copia coluna custo para coluna valor_custo
            //Conciliacao.CopiaColuna(pathPlanilha, numColunaOrigem, numColunaDestino, linhaInicio);

            int numColuna = 8;
            int numColunaProcura = 1;
            string rangeColuna = "h:h";
            string letraColuna = "h";
            //int linhaInicio = 2;
            rangeProcura = "a:a";

            // soma total SD1
            //Conciliacao.CalculaTotalSD1(pathPlanilha, numColuna, letraColuna, rangeColuna, linhaInicio, rangeProcura, numColunaProcura);
           
            //soma total SD2
            //Conciliacao.CalculaTotalSD2(pathPlanilha, numColuna, letraColuna, rangeColuna, linhaInicio, rangeProcura, numColunaProcura);


//            Conciliacao.CalculaDiferenca(pathPlanilha);
           
           // Conciliacao.SomaDebito();
           //Conciliacao.SomaCredito();
           // Conciliacao.CompareEntradaFilia01();
           //Conciliacao.CompareSaidaFilia01();

            //Conciliacao.TratamentoString(pathPlanilha, pathPlanilhaRazao);

            Conciliacao.lancamentosIncorretos(pathPlanilhaRazao, pathPlanilha);


            Console.Write("finished");
            Console.ReadLine();
        }
    }

}