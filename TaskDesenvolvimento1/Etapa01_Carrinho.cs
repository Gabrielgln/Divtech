using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Ajax.Utilities;
using PULSE.Models;
using PULSE.Util;
using Quartz;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Threading.Tasks;

namespace PULSE.SchedulerHelpers.Intemidia.Download
{
    
    public class Etapa01_Carrinho : IJob
    {
        private int dtRowCount = 0;
        private DataTable dtpedidos = null;

        private string strConnIntermidia = ConfigurationManager.ConnectionStrings["Intermidia_ConnString"].ConnectionString;
        private string strConnTess = ConfigurationManager.ConnectionStrings["PULSECONN"].ConnectionString;


        public Task Execute(IJobExecutionContext contexta)
        {
            string query = "";
            string tabelaOrigem = "";
            string tabelaDestino = "";
            DataTable dtPedidos = null;

            SqlTransaction objTransacao = null;

            try
            {
                //ServiceMail.Send($"Job {__tabelaDeOrigem} iniciado", $"{DateTime.Now}");
                using (var objConnIntermidia = new SqlConnection(strConnIntermidia))
                {
                    using (var objConnTess = new SqlConnection(strConnTess))
                    {
                        try
                        {
                            //Preparar conexão com banco de dados.
                            objConnTess.Open();
                            objConnIntermidia.Open();
                            objTransacao = objConnTess.BeginTransaction();

                            //Processar ESR_CARRINHO
                            tabelaOrigem = "ESR_CARRINHO";
                            tabelaDestino = "INTM_ESR_CARRINHO";
                            query = $"SELECT * FROM ESR_CARRINHO WITH(NOLOCK) WHERE " +
                                    $"CodigoSituacaoPedido = 3 and DataIntegracao is null";
                            dtPedidos = ExecutarBulkCopy(objConnIntermidia,
                                    objConnTess,
                                    tabelaOrigem,
                                    tabelaDestino,
                                    query,
                                    objTransacao);

                            dtpedidos = dtPedidos;
                            //objTransacao.Commit();
                            //Rollback é para fins de teste
                            objTransacao.Rollback();
                        }
                        catch (Exception ex)
                        {
                            if (objTransacao != null)
                                objTransacao.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceMail.SendAlert(
                $"Job '{nameof(Etapa01_Carrinho)}' encerrado com erro",
                $"Erro: '{ex.Message}'. {DateTime.Now}.",
                "suporte@divtech.com.br"
                );
            }
           
            return Task.CompletedTask;
        }

        public int getDtRowCount()
        {
            dtRowCount = dtpedidos.Rows.Count;
            return dtRowCount;
        }

        private DataTable ExecutarBulkCopy(
            SqlConnection objConnIntermidia,
            SqlConnection objConnTess,
            string tabelaOrigem,
            string tabelaDestino,
            string query,
            SqlTransaction objTransacao)
        {
            DataTable dt = new DataTable(tabelaOrigem);
            var cmd = new SqlCommand(query);
            cmd.CommandTimeout = TimeSpan.FromMinutes(600).Seconds;
            cmd.Connection = objConnIntermidia;
            var da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                // Copia os dados a partir do DataTable para a tabela SQL
                using (var bulkCopy = new SqlBulkCopy(objConnTess, SqlBulkCopyOptions.Default, objTransacao))
                {
                    // os nomes das colunas do DataTable usado correspondem aos nomes das colunas da tabela SQL,
                    // assim usei um laço foreach simples. No entanto, se os nomes das colunas
                    // não corresponderem, apenas passe qual nome do DataTable corresponde ao nome da coluna SQL 
                    // em ColumnMappings
                    foreach (DataColumn col in dt.Columns)
                    {
                        bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                    }
                    bulkCopy.BulkCopyTimeout = 600;
                    bulkCopy.DestinationTableName = $"{tabelaDestino}";
                    bulkCopy.WriteToServer(dt);
                }
            }
            return dt;
        }
    }
}

