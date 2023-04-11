using PULSE.Util;
using Quartz;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.Text;
using System.Threading.Tasks;

namespace PULSE.SchedulerHelpers.Intemidia.Download
{
    public class Etapa03_Monitoramento_IntegracaoCliente : IJob
    {
        private string strConnTess = ConfigurationManager.ConnectionStrings["PULSECONN"].ConnectionString;
        
       
        public Task Execute(IJobExecutionContext contexta)
        {
            string tabela = "SA1100";
            string emailDoSuporte = "suporte@divtech.com.br";
            DataTable dataTable = null;
            var dataDaAplicacao = "";
            int countRows = 0;
            
            try
            {
                using (var objConn = new SqlConnection(strConnTess))
                {
                    objConn.Open();
                    dataTable = getTable(tabela, objConn);

                    dataDaAplicacao = DateTime.Now.ToString("yyyyMMdd");
                    countRows = dataTable.Rows.Count;

                    for (int i = countRows-1; i >= 0; i--)
                    {
                        if (dataTable.Rows[i]["DtCadastro"].ToString() != dataDaAplicacao)
                        {
                            dataTable.Rows.RemoveAt(i);
                        }
                    }
                    if (dataTable.Rows.Count > 0)
                    {
                        SendEmail(dataTable, emailDoSuporte);   
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceMail.SendAlert(
                $"Job '{nameof(Etapa03_Monitoramento_IntegracaoCliente)}' encerrado com erro",
                $"Erro: '{ex.Message}'. {DateTime.Now}.",
                "suporte@divtech.com.br"
                );
            }
            return Task.CompletedTask;
        }
        private DataTable getTable(string tabela, SqlConnection objConnDestino)
        {
            DataTable dataTable = new DataTable(tabela);
            var dataDaAplicacao = DateTime.Now.ToString("yyyyMMdd");
            ini = Convert.ToDateTime(data_ini);

            var query = "SELECT (A1_DTCAD) DtCadastro, (A1_CGC) Cnpj, (A1_NOME) RazaoSocial, (A1_NREDUZ) NomeFantasia, (A1_EST) UF " +
                        $"FROM DADOSADV.dbo.{tabela} " +
                        "WHERE D_E_L_E_T_ = '' " +
                        "AND ISNULL(A1_VEND,'') <> '' " +
                        "AND A1_VEND <> '319' " +
                        "AND A1_FILIAL = '01' " +
                        "AND LEN(A1_CGC) = 14 " +
                        "AND A1_CGC <> '00000000000000' " +
                        "AND A1_CGC <> '99999999999999' " +
                        "AND A1_CGC <> '00000000000009' " +
                        "AND A1_CGC <> '00000000000002' " +
                        "AND A1_CGC <> '00000000000001' " +
                        "AND A1_COD <> '' " +
                        "AND A1_LOJA <> '' " +
                        "AND A1_STATUS = 'A' " +
                        "AND A1_NOME <> '' " +
                        "AND A1_END <> '' " +
                        "AND A1_MCOMPRA <> '' "+
                        "AND A1_PESSOA = 'J' " +
                        "AND A1_TIPO <> 'R'";

            var cmd = new SqlCommand(query, objConnDestino);
            cmd.CommandTimeout = TimeSpan.FromMinutes(600).Seconds;

            var dataAdapter = new SqlDataAdapter(cmd);
            dataAdapter.Fill(dataTable);

            return dataTable;
        }
        private void SendEmail(DataTable dataTable, string emailDoSuporte)
        {
            var mensagemDoSuporte = new StringBuilder();
            var assuntoDaMensagemSuporte = "";
            var dataDoEmail = "";
            //Informações
            mensagemDoSuporte.AppendLine($"<p>Destinatário: {emailDoSuporte}</p>");
            mensagemDoSuporte.AppendLine("<p>Categoria: Suporte</p>");
            mensagemDoSuporte.AppendLine("<p>Prioriade: Muito Alta</p>");
            //Definindo o cabeçalho da tabela da mensagem
            mensagemDoSuporte.AppendLine("<table style='float: left;' border='1' cellspacing='0' cellpadding='5'>");
            mensagemDoSuporte.AppendLine("<tr>");
            mensagemDoSuporte.AppendLine("<th>Data do cadastro</th>");
            mensagemDoSuporte.AppendLine("<th>Cnpj do Cliente</th>");
            mensagemDoSuporte.AppendLine("<th>Razão Social do Cliente</th>");
            mensagemDoSuporte.AppendLine("<th>Nome Fantasia do Cliente</th>");
            mensagemDoSuporte.AppendLine("<th>UF do Cliente</th>");
            mensagemDoSuporte.AppendLine("</tr>");
            //Passando as informações da tabela do banco de dados para a tabela da mensagem, linha por linha
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                mensagemDoSuporte.AppendLine("<tr>");
                mensagemDoSuporte.AppendLine($"<td>{dataTable.Rows[i]["DtCadastro"]}</td>");
                mensagemDoSuporte.AppendLine($"<td>{dataTable.Rows[i]["Cnpj"]}</td>");
                mensagemDoSuporte.AppendLine($"<td>{dataTable.Rows[i]["RazaoSocial"]}</td>");
                mensagemDoSuporte.AppendLine($"<td>{dataTable.Rows[i]["NomeFantasia"]}</td>");
                mensagemDoSuporte.AppendLine($"<td>{dataTable.Rows[i]["UF"]}</td>");
                mensagemDoSuporte.AppendLine("<tr>");
            }
            mensagemDoSuporte.AppendLine("</table>");


            //Definindo a data e hora do email atual
            dataDoEmail = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            //Definindo o assunto da mensagem para o suporte
            assuntoDaMensagemSuporte = $"Alerta: [{dataDoEmail}] Relatório de clientes com falha de integração";

            //Enviando o email com os parâmetros preenchidos
            //ServiceMail.SendAlert(assuntoDaMensagemSuporte, mensagemDoSuporte.ToString(), emailDoSuporte);

            //Limpar as mensagens para os próximos emails
            mensagemDoSuporte.Clear();
        }
    }
}