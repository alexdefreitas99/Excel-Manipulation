
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using classes;
using TestVai;

namespace ConciliacaoEstoque
{

    public class Filial
    {
        public string numeroFilial { get; set; }
        public int numeroNotaRazao { get; set; }
        public double debito { get; set; }
        public int linhaExcel { get; set; }
        public bool deletada { get; set; }
        public string addressFilial { get; set; }
    }

    public class Nota
    {
        public double custo { get; set; }
        public int numero { get; set; }
        public int numeroNotaEstoque { get; set; }
        public int linhaExcel { get; set; }
        public bool deletada { get; set; }
        public string addressNota { get; set; }
        public int rowNota { get; set; }
    }

    public class Conciliacao
    {
        string pathPlanilhaRazao = @"C:\planilhasEstoque\rz_todas_filias_1.xlsx";
        string pathPlanilhaConta = @"C:\planilhasEstoque\01_CONTA_CCL_CMPLETA.xlsx";
        string extension = "xlsx";

        public static void LimpaFormatacaoArquivoExcel()
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excel.Workbooks.Open(pathPlanilhaConta);
            excel.DisplayAlerts = false;

            try
            {
                for (int i = excel.ActiveWorkbook.Worksheets.Count; i > 0; i--)
                {
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveWorkbook.Worksheets[i];
                    worksheet.Cells.ClearFormats();
                }
            }
            finally
            {
                workbook.Save();
                workbook.Close();
                excel.Quit();
                excel.DisplayAlerts = true;
            }
        }

        public static void MakeUsableCopyFromCrashedExcelFile()
        {
            string tempPath = pathPlanilhaConta.Replace($".{extension}", $"_temp.{extension}");

            using (var package = new ExcelPackage(new FileInfo(pathPlanilhaConta)))
            {

                ExcelWorksheet workSheet = package.Workbook.Worksheets.ToList().First();

                ExcelWorksheet worksheetCopiado = package.Workbook.Worksheets.Copy(workSheet.Name, workSheet.Name + "_copy");

                using (var novoExcel = new ExcelPackage())
                {
                    var worksheet = novoExcel.Workbook.Worksheets.Add(worksheetCopiado.Name, worksheetCopiado);
                    novoExcel.SaveAs(new FileInfo(tempPath));
                }
            }

            File.Delete(pathPlanilhaConta);
            File.Move(tempPath, pathPlanilhaConta);
        }

        public static string DeletaLinhasPlanilha(string palavraChave, string rangeProcuraPalavra)
        {
            var planilha = new ExcelPackage(new FileInfo(pathPlanilhaConta));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            var start = planilhaAba1.Dimension.Start;
            var end = planilhaAba1.Dimension.End;
            string RangePalavraChave = start.ToString() + rangeProcuraPalavra + end.ToString();

            var LocalizaPalavraChave = from cell in planilhaAba1.Cells[rangeProcuraPalavra]
                                       where cell.Value.ToString() == palavraChave
                                       select cell.Start.Row;


            int? linhaPalavraChaveEncontrada = null;

            for (int row = 1; row <= end.Row; row++)
            {
                for (int x = 1; x < end.Column; x++)
                {
                    object cellValue = planilhaAba1.Cells[row, x].Value;

                    if (palavraChave.Equals(cellValue))
                    {
                        linhaPalavraChaveEncontrada = row;
                        break;
                    }
                }

                if (linhaPalavraChaveEncontrada != null)
                    break;
            }


            //apagar de ultimaLinha pra baixo
            if (linhaPalavraChaveEncontrada != null)
            {
                var linha = planilhaAba1.Workbook.Worksheets.FirstOrDefault();
                linha.DeleteRow(linhaPalavraChaveEncontrada.Value + 1, end.Row);
            }

            planilha.Save();

            return "";

        }

        public static bool LinhasValorDuplicado(string pathPlanilha, string rangeProcura, int linhaInicio)
        {
            var planilha = new ExcelPackage(new FileInfo(pathPlanilha));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            var start = planilhaAba1.Dimension.Start;

            List<Nota> notas = new List<Nota>();

            var end = planilhaAba1.Dimension.End;

            for (int i = linhaInicio; i < end.Row; i++)
            {

                double valorCusto = 1;
                double.TryParse(planilhaAba1.Cells[i, 9].Value.ToString(), out valorCusto);

                int numero = 0;
                int.TryParse(planilhaAba1.Cells[i, 19].Value.ToString(), out numero);

                Nota nota = new Nota();
                nota.custo = valorCusto;
                nota.numero = numero;
                nota.linhaExcel = i;
                nota.deletada = false;

                notas.Add(nota);
            }

            var duplicadosAgrupados = notas.GroupBy(x => new { x.custo, x.numero });
            foreach (var duplicados in duplicadosAgrupados)
	        {
               if (duplicados.Count() > 1)
	            {
                    var i = duplicados.Count();
		            foreach (var deletar in duplicados)
	                {

                        var linha = planilhaAba1.Workbook.Worksheets.FirstOrDefault();
                        linha.DeleteRow(deletar.linhaExcel);
                        
                        i -= 1;

                        if (i==1) 
                            break;
	                }
	            }
	        }
            
            planilha.Save();

            return true;
        }        

        public static void SomaDebito(){        
          var planilha = new ExcelPackage(new FileInfo(pathPlanilhaRazao));
          ExcelWorksheet planilhaRazaoAba1 = planilha.Workbook.Worksheets.FirstOrDefault();
           
            double resultado1 = 0;
            for (int i = 14; i < planilhaRazaoAba1.Dimension.End.Row; i++)
            {
                if(planilhaRazaoAba1.Cells[i, 5].Value!=null){
                    if(planilhaRazaoAba1.Cells[i, 5].Value.Equals("01")){
                            resultado1 +=  Double.Parse(planilhaRazaoAba1.Cells[i,9].First().Value.ToString());
                    }
                }
            } 
            planilhaRazaoAba1.Cells["I"+planilhaRazaoAba1.Dimension.End.Row].Clear();
            planilhaRazaoAba1.Cells["I"+planilhaRazaoAba1.Dimension.End.Row].Value = resultado1;
            planilha.Save();
        }

        public static void SomaCredito(){
          //Soma credito conforme a filial
          var planilha = new ExcelPackage(new FileInfo(pathPlanilhaRazao));
          ExcelWorksheet planilhaRazaoAba1 = planilha.Workbook.Worksheets.FirstOrDefault();
            double resultado1 = 0;
            for (int i = 14; i < planilhaRazaoAba1.Dimension.End.Row; i++)
            {
                if(planilhaRazaoAba1.Cells[i, 5].Value!=null){
                    if(planilhaRazaoAba1.Cells[i, 5].Value.Equals("01")){
                            resultado1 +=  Double.Parse(planilhaRazaoAba1.Cells[i, 10].First().Value.ToString());
                    }
                }
            } 
            planilhaRazaoAba1.Cells["J"+planilhaRazaoAba1.Dimension.End.Row].Clear();
            planilhaRazaoAba1.Cells["J"+planilhaRazaoAba1.Dimension.End.Row].Value = resultado1;
            planilha.Save();
        }

        public static void CompareEntradaFilia01(){
          var planilha = new ExcelPackage(new FileInfo(pathPlanilhaRazao));
          ExcelWorksheet planilhaRazaoAba1 = planilha.Workbook.Worksheets.FirstOrDefault();
             
          var planilhaConta = new ExcelPackage(new FileInfo(pathPlanilhaConta));
          ExcelWorksheet planilhaContaAba1 = planilhaConta.Workbook.Worksheets.FirstOrDefault();

            string debitoTotal  = "";
            string sd1Total = "";
            string addressParaPintar = "";

            for (int i = 1; i <=  planilhaContaAba1.Dimension.End.Row; i++)
            {
                if(planilhaContaAba1.Cells[i, 1].Value!=null)
                {
                    if(planilhaContaAba1.Cells[i, 1].Value.Equals("SD1-TOTAL"))
                    {
                        sd1Total = planilhaContaAba1.Cells[i, 8].Value.ToString();
                        addressParaPintar = planilhaContaAba1.Cells[i, 8].Address;
                        break;
                    }
                }
            }
            
            if(planilhaRazaoAba1.Cells["I"+planilhaRazaoAba1.Dimension.End.Row].Value.ToString()  == sd1Total){
                    Console.WriteLine("acertou");
            }
            else{
                planilhaContaAba1.Cells[addressParaPintar].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                planilhaContaAba1.Cells[addressParaPintar].Style.Fill.BackgroundColor.SetColor(Color.Red);
                planilhaContaAba1.Cells[addressParaPintar].Style.Font.Color.SetColor(Color.Black);
                planilhaConta.Save();

                planilhaRazaoAba1.Cells["I"+planilhaRazaoAba1.Dimension.End.Row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                planilhaRazaoAba1.Cells["I"+planilhaRazaoAba1.Dimension.End.Row].Style.Fill.BackgroundColor.SetColor(Color.Red);
                planilhaRazaoAba1.Cells["I"+planilhaRazaoAba1.Dimension.End.Row].Style.Font.Color.SetColor(Color.Black);
                TratamentoString(pathPlanilhaConta, pathPlanilhaRazao);
                planilha.Save();
             }            
        }

        public static void lancamentosIncorretos()
        {
          var planilha = new ExcelPackage(new FileInfo(pathPlanilhaRazao));
          ExcelWorksheet planilhaRazaoAba1 = planilha.Workbook.Worksheets.FirstOrDefault();
          var planilhaConta = new ExcelPackage(new FileInfo(pathPlanilhaConta));
          ExcelWorksheet planilhaAba1 = planilhaConta.Workbook.Worksheets.FirstOrDefault();
            string pathSaveTxt = @"C:\planilhasEstoque\teste.txt";
            List<Nota> notas = new List<Nota>();
            List<Filial> filiais = new List<Filial>();

            for (int i = 2; i < planilhaAba1.Dimension.End.Row; i++)
            {
                if (planilhaAba1.Cells[i, 19].Value != null)
	            {
                    if (planilhaAba1.Cells[i, 19].Value != "")
	                {
                        Nota nota = new Nota();
                        int notasEstoque;
                        double custo;
                        int.TryParse(planilhaAba1.Cells[i, 19].Value.ToString(), out notasEstoque);
                        double.TryParse(planilhaAba1.Cells[i, 8].Value.ToString(), out custo);
                        nota.numeroNotaEstoque = notasEstoque;
                        nota.custo = custo;
                        nota.addressNota = planilhaAba1.Cells[i, 19].Address;
                        nota.rowNota = planilhaAba1.Cells[i, 19].End.Row;
                        notas.Add(nota);
	                }
	            }
            }
            
            for (int i = 14; i < planilhaRazaoAba1.Dimension.End.Row; i++)
            {
                if (planilhaRazaoAba1.Cells[i, 3].Value != null)
	            {
                    if (planilhaRazaoAba1.Cells[i, 3].Value != "")
	                {
                        string[] numbers = Regex.Split(planilhaRazaoAba1.Cells[i, 3].Value.ToString(), @"\D+");
                        if (!string.IsNullOrEmpty(numbers[1]))
                        {
                            Filial historico = new Filial();
                            int numeroRazao = int.Parse(numbers[1]);
                            double debito;
                            double.TryParse(planilhaRazaoAba1.Cells[i, 9].Value.ToString(), out debito);
                            historico.numeroNotaRazao = numeroRazao;
                            historico.debito = debito;
                            historico.addressFilial = planilhaRazaoAba1.Cells[i, 3].Address;
                            filiais.Add(historico);  
                        }
	                } 
	            }   
            }

            List<Nota> notasDuplicadas = new List<Nota>();
            string result = "";
            
            for (int i = 0; i < notas.Count; i++ )
			{
                for (int j = 0; j < filiais.Count; j++)
			    {
                    if (notas[i].numeroNotaEstoque == filiais[j].numeroNotaRazao)
	                {
                        if (notas[i].custo != filiais[j].debito)
	                    {
                            Console.WriteLine(notas[i].numero);
                            planilhaAba1.Cells["A"+notas[i].rowNota+":AE"+notas[i].rowNota].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            planilhaAba1.Cells["A"+notas[i].rowNota+":AE"+notas[i].rowNota].Style.Fill.BackgroundColor.SetColor(Color.Black);
                            planilhaAba1.Cells["A"+notas[i].rowNota+":AE"+notas[i].rowNota].Style.Font.Color.SetColor(Color.White);
                            result += "Custo: "+notas[i].custo+"Nota: "+notas[i].numero+";"; 
	                    }
	                }
			    }
			}
            planilhaConta.Save();           
            File.WriteAllText(pathSaveTxt, result);

        }

        public static void CompareSaidaFilia01(){
            
          var planilha = new ExcelPackage(new FileInfo(pathPlanilhaRazao));
          ExcelWorksheet planilhaRazaoAba1 = planilha.Workbook.Worksheets.FirstOrDefault();
          var planilhaConta = new ExcelPackage(new FileInfo(pathPlanilhaConta));
          ExcelWorksheet planilhaContaAba1 = planilhaConta.Workbook.Worksheets.FirstOrDefault();

            string creditoTotal  = "";
            string sd2Total = "";
            string addressParaPintar = "";

            for (int i = 1; i <=  planilhaContaAba1.Dimension.End.Row; i++)
            {
                if(planilhaContaAba1.Cells[i, 1].Value!=null)
                {
                    if(planilhaContaAba1.Cells[i, 1].Value.Equals("SD2-TOTAL"))
                    {
                        sd2Total = planilhaContaAba1.Cells[i, 8].Value.ToString();
                        addressParaPintar = planilhaContaAba1.Cells[i, 8].Address;
                        break;
                    }
                }
            }
            
            if(planilhaRazaoAba1.Cells["J"+planilhaRazaoAba1.Dimension.End.Row].Value.ToString()  == sd2Total){
                    Console.WriteLine("acertou");
            }
            else{
                planilhaContaAba1.Cells[addressParaPintar].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                planilhaContaAba1.Cells[addressParaPintar].Style.Fill.BackgroundColor.SetColor(Color.Red);
                planilhaContaAba1.Cells[addressParaPintar].Style.Font.Color.SetColor(Color.Black);
                planilhaConta.Save();

                planilhaRazaoAba1.Cells["J"+planilhaRazaoAba1.Dimension.End.Row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                planilhaRazaoAba1.Cells["J"+planilhaRazaoAba1.Dimension.End.Row].Style.Fill.BackgroundColor.SetColor(Color.Red);
                planilhaRazaoAba1.Cells["J"+planilhaRazaoAba1.Dimension.End.Row].Style.Font.Color.SetColor(Color.Black);
                planilha.Save();
             }            
        }

        public static string CopiaColuna(int numColunaOrigem, int numColunaDestino, int linhaInicio)
        {
            var planilha = new ExcelPackage(new FileInfo(pathPlanilhaConta));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            var start = planilhaAba1.Dimension.Start;
            var end = planilhaAba1.Dimension.End;

            planilhaAba1.Cells[linhaInicio, numColunaOrigem, end.Row, numColunaOrigem].Copy(planilhaAba1.Cells[linhaInicio, numColunaDestino, end.Row, numColunaDestino]);

            planilha.Save();

            return "";
        }

        public static string CalculaTotalSD2(int numColuna, string letraColuna, string rangeColuna, int linhaInicio, string rangeProcura, int numColunaProcura)
        {
            var planilha = new ExcelPackage(new FileInfo(pathPlanilhaConta));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            var start = planilhaAba1.Dimension.Start;
            var end = planilhaAba1.Dimension.End;
            var palavraChave = "SD2";
            var palavraChave2 = "SD2-TOTAL";
           // int primeiraLinha = 0;
            int ultimaLinha = 0;
            var celulaSelecionar = "";
            var primeiraCelula = "";
            var ultimaCelula = "";

            // Localiza primeira palavra chave2 SD2
            IEnumerable<ExcelRangeBase> LocalizaPalavraChave = from cell in planilhaAba1.Cells[rangeProcura]
                                                               where cell.Value.ToString() == palavraChave
                                                               select cell;
            var excelRangeBase = LocalizaPalavraChave.First();
            primeiraCelula = excelRangeBase.Address;
            var primeiraLinha = excelRangeBase.End.Row;
            
            // Localiza primeira palavra chave SD2-TOTAL
            IEnumerable<ExcelRangeBase> LocalizaPalavraChave2 = from cell in planilhaAba1.Cells[rangeProcura]
                                                                where cell.Value.ToString() == palavraChave2
                                                                select cell;
            excelRangeBase = LocalizaPalavraChave2.First();
            ultimaCelula = excelRangeBase.Address;
            ultimaLinha = excelRangeBase.End.Row;

            //Limpar célula para aplicar fórmula.

            var selecionaCelula = planilhaAba1.Cells[ultimaLinha, numColuna];
            celulaSelecionar = selecionaCelula.ToString();

            planilhaAba1.Cells[celulaSelecionar].Clear();
            ultimaLinha = ultimaLinha - 1;
            ultimaCelula = letraColuna + ultimaLinha;

            //planilhaAba1.Cells[celulaSelecionar].Formula = "=SUM(" + primeiraCelula + ":" + ultimaCelula + ")";
            
            double resultado = 0;
            string coluna = "h";
            string celula;
            for (int i = primeiraLinha; i <= ultimaLinha; i++)
			{
                celula = coluna + i;
                resultado += Double.Parse(planilhaAba1.Cells[celula].Value.ToString());
			}

            planilhaAba1.Cells[celulaSelecionar].Value = resultado+"";
            
            planilha.Save();


            return "";
        }

        public static string CalculaTotalSD1(int numColuna, string letraColuna, string rangeColuna, int linhaInicio, string rangeProcura, int numColunaProcura)
        {
            var planilha = new ExcelPackage(new FileInfo(pathPlanilhaConta));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            var start = planilhaAba1.Dimension.Start;
            var end = planilhaAba1.Dimension.End;
            var palavraChave = "SD1";
            var palavraChave2 = "SD1-TOTAL";
            int primeiraLinha = 0;
            int ultimaLinha = 0;
            var celulaSelecionar = "";
            var primeiraCelula = "";
            var ultimaCelula = "";

            // Localiza primeira palavrachave SD1
            IEnumerable<ExcelRangeBase> LocalizaPalavraChave = from cell in planilhaAba1.Cells[rangeProcura]
                                                               where cell.Value.ToString() == palavraChave
                                                               select cell;
            var excelRangeBase = LocalizaPalavraChave.First();
            primeiraCelula = excelRangeBase.Address;
    

            // Localiza primeira palavrachave2 SD1-TOTAL
            IEnumerable<ExcelRangeBase> LocalizaPalavraChave2 = from cell in planilhaAba1.Cells[rangeProcura]
                                                                where cell.Value.ToString() == palavraChave2
                                                                select cell;
            excelRangeBase = LocalizaPalavraChave2.First();
            ultimaCelula = excelRangeBase.Address;
            ultimaLinha = excelRangeBase.End.Row;

            //Limpar célula para aplicar fórmula.

            var selecionaCelula = planilhaAba1.Cells[ultimaLinha, numColuna];
            celulaSelecionar = selecionaCelula.ToString();

            planilhaAba1.Cells[celulaSelecionar].Clear();
            ultimaLinha = ultimaLinha - 1;
            ultimaCelula = letraColuna + ultimaLinha;

            //planilhaAba1.Cells[celulaSelecionar].Formula = "=SUM(" + primeiraCelula + ":" + ultimaCelula + ")";

            double resultado = 0;
            string coluna = "h";
            string celula;
            for (int i = 2; i <= ultimaLinha; i++)
			{
                celula = coluna + i;
                resultado += Double.Parse(planilhaAba1.Cells[celula].Value.ToString());
			}

            planilhaAba1.Cells[celulaSelecionar].Value = resultado+"";
            planilha.Save();

            return "";
        }

        public static string CalculaDiferenca()
        {
            var planilha = new ExcelPackage(new FileInfo(pathPlanilhaConta));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            var start = planilhaAba1.Dimension.Start;
            var end = planilhaAba1.Dimension.End;
            
            var colunaValorCusto = "H";
            var ColunaCusto = "I";
            var colunaDiferenca = "J";
            double v1;
            double v2;
            double result;
            for (int linhaFor = 2; linhaFor <= end.Row; linhaFor++)
            {                            
                var celulaValorCusto = colunaValorCusto + linhaFor;
                var celulaCusto = ColunaCusto + linhaFor;
                var celulaDiferenca = colunaDiferenca + linhaFor;
                planilhaAba1.Cells[celulaDiferenca].Clear();       
                v1 = Double.Parse(planilhaAba1.Cells[celulaValorCusto].Value.ToString());
                v2 = Double.Parse(planilhaAba1.Cells[celulaCusto].Value.ToString());
                result = v1-v2;
                planilhaAba1.Cells[celulaDiferenca].Value = result;
            }          
            
             for (int i = 2; i <= end.Row; i++) {
                var celulasPintar = colunaDiferenca + i;
                //planilhaAba1.Cells[celulasPintar].First().Calculate();
                var range = planilhaAba1.Cells[celulasPintar];
                //var t = planilhaAba1.Cells[celulasPintar].First().Value.ToString();
                planilhaAba1.Cells[celulasPintar].Calculate();
                if (planilhaAba1.Cells[celulasPintar].Value.ToString() != ("0")) {
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    range.Style.Font.Color.SetColor(Color.Black);
                }
             }
              

            planilha.Save();
            
            return "";
                
        }

        public static void ExtraiFilial(int linhaInicio)
        {
           var planilha = new ExcelPackage(new FileInfo(pathPlanilhaRazao));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            for (int i = 14; i <=  planilhaAba1.Dimension.End.Row; i++)
            {
                if(planilhaAba1.Cells[i, 5].Value!=null)
                {
                    if(planilhaAba1.Cells[i, 5].Value.Equals("02") || planilhaAba1.Cells[i, 5].Value.Equals("05"))
                    {
                        planilhaAba1.DeleteRow(i);
                        i -= 1;
                    }
                }
            }
            planilha.SaveAs(new FileInfo(@"C:\planilhasEstoque\rz_todas_filias_"+01+".xlsx"));

        }
            
        public static void TratamentoString()
        {
            var planilha = new ExcelPackage(new FileInfo(pathPlanilhaConta));
            ExcelWorksheet planilhaAba1 = planilha.Workbook.Worksheets.FirstOrDefault();

            var planilhaRazao = new ExcelPackage(new FileInfo(pathPlanilhaRazao));
            ExcelWorksheet planilhaRazaoAba1 = planilhaRazao.Workbook.Worksheets.FirstOrDefault();

            List<Nota> notas = new List<Nota>();
            List<Filial> filiais = new List<Filial>();

            for (int i = 2; i < planilhaAba1.Dimension.End.Row; i++)
            {
                if (planilhaAba1.Cells[i, 19].Value != null)
	            {
                    if (planilhaAba1.Cells[i, 19].Value != "")
	                {
                        Nota nota = new Nota();
                        int notasEstoque;
                        int.TryParse(planilhaAba1.Cells[i, 19].Value.ToString(), out notasEstoque);
                        nota.numeroNotaEstoque = notasEstoque;
                        nota.addressNota = planilhaAba1.Cells[i, 19].Address;
                        nota.rowNota = planilhaAba1.Cells[i, 19].End.Row;
                        notas.Add(nota);
	                }
	            }
            }

            for (int i = 1; i < planilhaRazaoAba1.Dimension.End.Row; i++)
            {
                if (planilhaRazaoAba1.Cells[i, 3].Value != null)
	            {
                    if (planilhaRazaoAba1.Cells[i, 3].Value != "")
	                {
                        string[] numbers = Regex.Split(planilhaRazaoAba1.Cells[i, 3].Value.ToString(), @"\D+");
                        if (!string.IsNullOrEmpty(numbers[1]))
                        {
                            Filial historico = new Filial();
                            int numeroRazao = int.Parse(numbers[1]);
                            historico.numeroNotaRazao = numeroRazao;
                            historico.addressFilial = planilhaRazaoAba1.Cells[i, 3].Address;
                            filiais.Add(historico);                           
                        }
	                }
	            }
            }            
            
            for (int i = 0; i < notas.ToArray().Length; i++ )
			{
                for (int j = 0; j < filiais.ToArray().Length; j++)
			    {
                    if (notas[i].numeroNotaEstoque == filiais[j].numeroNotaRazao)
	                {
                        Console.WriteLine("ok");
	                }else{
                        planilhaAba1.Cells["A"+notas[i].rowNota+":AE"+notas[i].rowNota].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        planilhaAba1.Cells["A"+notas[i].rowNota+":AE"+notas[i].rowNota].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                        planilhaAba1.Cells["A"+notas[i].rowNota+":AE"+notas[i].rowNota].Style.Font.Color.SetColor(Color.Black);  
                    }
			    }
			}
            planilha.Save();
        }
    }
}

