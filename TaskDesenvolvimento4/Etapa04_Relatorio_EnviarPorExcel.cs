using Aspose.Cells;
using DocumentFormat.OpenXml.Office2013.Excel;
using Microsoft.Ajax.Utilities;
using PULSE.Util;
using Quartz;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Compilation;

namespace PULSE.SchedulerHelpers.Intemidia.Download
{
    public class Etapa04_Relatorio_EnviarPorEmail : IJob
    {
        private string strConnTess = ConfigurationManager.ConnectionStrings["PULSECONN"].ConnectionString;
        private string strConnWiseSale = ConfigurationManager.ConnectionStrings["WiseSale"].ConnectionString;

        public Task Execute(IJobExecutionContext contexta)
        {
            string RDV_INTERMIDIA_V1 = "VW_INTM_RELATORIO_COMERCIAL_V1";
            string RDV_INTERMIDIA_POR_MODELO = "VW_INTM_RELATORIO_COMERCIAL_POR_MODELO";
            string RDV_WISESALE = "RELATORIO_DIARIO_DE_PEDIDOS_ACUMULADO";

            DataTable dataTableV1 = null;
            DataTable dataTableModelo = null;
            DataTable dataTableWiseSales = null;

            try
            {
                using (var objConnTess = new SqlConnection(strConnTess))
                {
                    using (var objConnWiseSales = new SqlConnection(strConnWiseSale))
                    {
                        objConnTess.Open();
                        objConnWiseSales.Open();

                        dataTableV1 = getTablePulse(RDV_INTERMIDIA_V1, objConnTess);
                        dataTableModelo = getTablePulse(RDV_INTERMIDIA_POR_MODELO, objConnTess);
                        dataTableWiseSales = getTableWiseSales(RDV_WISESALE, objConnWiseSales);

                        CreateExcelDocument(dataTableV1);
                        CreateExcelDocument(dataTableModelo);
                        CreateExcelDocument(dataTableWiseSales);
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceMail.SendAlert(
                $"Job '{nameof(Etapa04_Relatorio_EnviarPorEmail)}' encerrado com erro",
                $"Erro: '{ex.Message}'. {DateTime.Now}.",
                "suporte@divtech.com.br"
                );
            }
            return Task.CompletedTask;
        }

        private DataTable getTablePulse(string tabela, SqlConnection objConnDestino)
        {
            DataTable dataTable = new DataTable(tabela);

            var query = "SELECT * " +
                        $"FROM PULSEBD.dbo.{tabela}";

            var cmd = new SqlCommand(query, objConnDestino);
            cmd.CommandTimeout = TimeSpan.FromMinutes(600).Seconds;

            var dataAdapter = new SqlDataAdapter(cmd);
            dataAdapter.Fill(dataTable);

            return dataTable;
        }

        private DataTable getTableWiseSales(string tabela, SqlConnection objConnDestino)
        {
            DataTable dataTable = new DataTable(tabela);

            var query = "SELECT * " +
                        $"FROM WISESALE.dbo.{tabela}";

            var cmd = new SqlCommand(query, objConnDestino);
            cmd.CommandTimeout = TimeSpan.FromMinutes(600).Seconds;

            var dataAdapter = new SqlDataAdapter(cmd);
            dataAdapter.Fill(dataTable);

            return dataTable;
        }

        private void CreateExcelDocument(DataTable dataTable)
        {
            try
            {
                using (var workbook = new Workbook())
                {
                    //Criando o alfabeto do excel - (A:Z - AA:ZZ...)
                    string[] Alphabet = new string[dataTable.Columns.Count];
                    for (int i = 0; i < Alphabet.Length; i++)
                    {
                        if (i < 26)
                        {
                            Alphabet[i] = ((char)('A' + i)).ToString();
                        }
                        else
                        {
                            char letter1 = (char)('A' + (i / 26) - 1);
                            char letter2 = (char)('A' + (i % 26));
                            Alphabet[i] = $"{letter1}{letter2}";
                        }
                    }
                    
                    //Criando a planilha
                    var worksheet = workbook.Worksheets.Add("Planilha1");
                    for (int i = 0; i < Alphabet.Length; i++)
                    {
                        //Pegando o titulo de cada coluna
                        DataColumn dtTitle = dataTable.Columns[i];
                        string dtTitleString = dtTitle.ColumnName;

                        for (int j = 0; j < dataTable.Rows.Count; j++)
                        {
                            //Formatando o valor do celula
                            string cellString = Alphabet[i].ToString();
                            int cellInt = j + 1;
                            string cellResult = cellString + cellInt.ToString();
                            //Inserindo valor da celula formatada
                            worksheet.Cells[cellResult].PutValue(dataTable.Rows[j][dtTitleString]);
                        }
                    }

                    //Formando o nome do arquivo e diretorio
                    var dateTime = DateTime.Now.ToString("yyyyMMddHHmm");
                    var fileName = $"RDV_{dateTime}.xlsx";
                    var user = Environment.UserName;

                    //Salvando o arquivo no diretorio de quem chamou a função
                    workbook.Save(@"C:\\Users\" + user + "\\Downloads\\" + fileName, SaveFormat.Xlsx);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}