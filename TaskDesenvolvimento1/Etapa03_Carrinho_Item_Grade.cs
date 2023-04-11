using PULSE.Util;
using Quartz;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PULSE.SchedulerHelpers.Intemidia.Download
{
    public class Etapa03_Carrinho_Item_Grade : IJob
    {
        private string strConnIntermidia = ConfigurationManager.ConnectionStrings["Intermidia_ConnString"].ConnectionString;
        private string strConnTess = ConfigurationManager.ConnectionStrings["PULSECONN"].ConnectionString;

        private Etapa01_Carrinho etapa01_Carrinho = new Etapa01_Carrinho();

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

                            if(etapa01_Carrinho.getDtRowCount() > 0) { 
                                //Processar ESR_CARRINHO
                                tabelaOrigem = "ESR_CARRINHO_ITEM_GRADE";
                                tabelaDestino = "INTM_ESR_CARRINHO_ITEM_GRADE";
                                query = "SELECT ecig.* FROM " +
                                    " MVW_Kenner.dbo.ESR_CARRINHO car WITH(NOLOCK)" +
                                    " inner join MVW_Kenner.dbo.ESR_CARRINHO_ITEM eci WITH(NOLOCK) on" +
                                    " car.CodigoCarrinho = eci.CodigoCarrinho" +
                                    " inner join MVW_Kenner.dbo.ESR_CARRINHO_ITEM_GRADE ecig WITH(NOLOCK) on" +
                                    " ecig.CodigoCarrinho = car.CodigoCarrinho" +
                                    " and ecig.SeqItem = eci.SeqItem" +
                                    " WHERE" +
                                    " car.CodigoSituacaoPedido = 3" +
                                    " and car.DataIntegracao is null";

                                dtPedidos = ExecutarBulkCopy(objConnIntermidia,
                                    objConnTess,
                                    tabelaOrigem,
                                    tabelaDestino,
                                    query,
                                    objTransacao);

                                //Confirmar recebimento do pedido
                                ConfirmarRecebimentoDoPedido(objConnIntermidia, dtPedidos);

                                //objTransacao.Commit();
                                objTransacao.Rollback();
                            }
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
                $"Job '{nameof(Etapa03_Carrinho_Item_Grade)}' encerrado com erro",
                $"Erro: '{ex.Message}'. {DateTime.Now}.",
                "suporte@divtech.com.br"
                );
            }
            return Task.CompletedTask;
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
        private void ConfirmarRecebimentoDoPedido(SqlConnection objConnIntermidia, DataTable dt)
        {
            var codigoCarrinho = "";
            var scriptSql = new System.Text.StringBuilder();
            var cmd = new SqlCommand("", objConnIntermidia);
            //Preparar SQL
            foreach (DataRow r in dt.Rows)
            {
                codigoCarrinho = r[nameof(codigoCarrinho)].ToString();
                scriptSql.AppendLine($"UPDATE {dt.TableName} " +
                    $"SET dataIntegracao = GETDATE() " +
                    $"WHERE {nameof(codigoCarrinho)} = '{codigoCarrinho}'; ");
            }
            cmd.CommandText = scriptSql.ToString();
            cmd.ExecuteNonQuery();
        }
    }
}
