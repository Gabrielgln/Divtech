using DocumentFormat.OpenXml.Office2010.Drawing;
using PULSE.Util;
using Quartz;
using System;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PULSE.SchedulerHelpers.Intemidia.Upload
{
    public class Etapa03_ClienteMarca : IJob
    {
        private string strConnIntermidia = ConfigurationManager.ConnectionStrings["IntermidiaSync_ConnString"].ConnectionString;
        private string strConnTess = ConfigurationManager.ConnectionStrings["PULSECONNSISCOMERCIAL"].ConnectionString;
        

        private string brandType;

        public void setBrandType(string brandType)
        {
            this.brandType = brandType;
        }


        public Task Execute(IJobExecutionContext contexta)
        {

            string query = "";
            string tabelaDestino = "CLIENTE_MARCA";
            string tabelaOrigem = $"VW_INTM_CLIENTE_MARCA";
            DataTable dt = null;
            
            try
            {
                //ServiceMail.Send($"Job {__tabelaDeOrigem} iniciado", $"{DateTime.Now}");
                using (var objConnIntermidia = new SqlConnection(strConnIntermidia))
                {
                    using (var objConnTess = new SqlConnection(strConnTess))
                    {
                        objConnTess.Open();
                        objConnIntermidia.Open();

                        //Iniciar cópia
                        query = $"SELECT * FROM {tabelaOrigem} WHERE CodigoMarca = '{brandType}' ";
                        //dt = ExecutarBulkCopy(objConnIntermidia, objConnTess, tabelaOrigem, tabelaDestino, query);
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceMail.SendAlert(
                    $"Job '{nameof(Etapa03_ClienteMarca)}' encerrado com erro",
                    $"Erro: '{ex.Message}'. {DateTime.Now}.",
                    "suporte@divtech.com.br"
                    );
            }
            finally
            {
                //ServiceMail.Send($"Job {__tabelaDeOrigem} encerrado", $"{DateTime.Now}");
            }

            Task.Delay(TimeSpan.FromSeconds(10));
            return Task.CompletedTask;
        }

        private DataTable ExecutarBulkCopy(
             SqlConnection objConnIntermidia,
             SqlConnection objConnTess,
             string tabelaOrigem,
             string tabelaDestino,
             string query)
        {
            DataTable dt = new DataTable(tabelaOrigem);
            var cmd = new SqlCommand(query);
            cmd.CommandTimeout = TimeSpan.FromMinutes(600).Seconds;
            cmd.Connection = objConnTess;
            var da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            //Limpar tabela antes de copiar os dados
            (new SqlCommand($"TRUNCATE TABLE {tabelaDestino}", objConnIntermidia)).ExecuteNonQuery();
            // Copia os dados a partir do DataTable para a tabela SQL
            using (var bulkCopy = new SqlBulkCopy(objConnIntermidia))
            {
                // os nomes das colunas do DataTable usado correspondem aos nomes das colunas da tabela SQL,
                // assim usei um laço foreach simples. No entanto, se os nomes das colunas
                // não corresponderem, apenas passe qual nome do DataTable corresponde ao nome da coluna SQL 
                // em ColumnMappings
                //foreach (DataColumn col in dt.Columns)
                //{
                //    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                //}
                bulkCopy.BulkCopyTimeout = 600;
                bulkCopy.DestinationTableName = $"{tabelaDestino}";
                bulkCopy.WriteToServer(dt);
            }
            return dt;
        }
    }
}