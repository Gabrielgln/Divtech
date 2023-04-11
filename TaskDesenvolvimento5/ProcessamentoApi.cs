using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using ProcessamentoRoboClasses.ApiSistemaComercial2.Models;
using ProcessamentoRoboClasses.ApiSistemaComercial2.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reactive.Linq;
using System.Reflection;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProcessamentoRoboClasses
{
    public class ProcessaApi : AcessoDados.Database
    {
        private static int total, processado, updated, entregue;
        private static DateTime inicio, termino;

        public static async Task processaNfsAsync()
        {
            var notaFiscal = "";
            var dataEmissao = DateTime.Now.AddMonths(-3).ToString("yyyyMMdd");
            inicio = DateTime.Now;
            termino = default(DateTime);
            try
            {
                AcessoDados.AcessaDadosSQL.DADOSADV = abrirConexao("REMOTO_DADOSADV");
                AcessoDados.AcessaDadosSQL.EXTRANET = abrirConexao("REMOTO_EXTRANET");
                DataTable dt = AcessoDados.AcessaDadosSQL.requisitaProcessamentoNfs(dataEmissao);
                List<Task> tasks = new List<Task>();

                total = dt.Rows.Count;
                processado = 0;
                updated = 0;
                entregue = 0;

                foreach (DataRow item in dt.Rows)
                {
                    notaFiscal = item["NOTA"].ToString();
                    ocorrencia objOcorrencia = new ocorrencia();
                    objOcorrencia.NumNf = item["NOTA"].ToString();
                    objOcorrencia.FilialERP = "01";
                    objOcorrencia.LojaClienteERP = item["LOJA_CLI"].ToString();
                    objOcorrencia.Obs = "";
                    objOcorrencia.DataLog = DateTime.Now.ToString();
                    objOcorrencia.CodClienteERP = item["CLIENTE"].ToString();
                    objOcorrencia.CodTransportadora = item["COD_TRANSP"].ToString();
                    objOcorrencia.EmissaoNf = item["EMISSAO"].ToString();

                    //Iniciando o objeto que faz a comunicação com a API
                    NotaFiscalServices notaFiscalServices = new NotaFiscalServices();
                    //Variavel que irá receber a lista que a API retornar
                    var notaFiscalApi = await notaFiscalServices.BuscarNotaFiscalPorNumero(int.Parse(objOcorrencia.NumNf));
                    
                    //var notaFiscalApi = await notaFiscalServices.BuscarNotaFiscalPorNumero(1082082);

                    //Verificação depois que a variavel receber a lista da API
                    if (notaFiscalApi.Count == 0 || notaFiscalApi == null)
                    {
                        continue;
                    }

                    bool processaBD = true;
                    //Iniciando um objeto de validação de dados
                    AcessoDados.Utils util = new AcessoDados.Utils();

                    //Variaveis que irão alimentar o Protheus
                    string previsaoEntregaProtheus = "";
                    string dataEntregaProteus = "";
                    string dataCancelamentoProtheus = "";
                   
                    foreach (var i in notaFiscalApi)
                    {
                        //Passando na lista de timelines
                        var timelineList = i.Timeline;
                        foreach(var y in timelineList)
                        {
                            //Verificando se tem alguma timeline de 'CANCELAMENTO' para atribuir a data de cancelamento
                            if(y.Descricao == "Pendente de Devolução" || y.Descricao == "Pendente de Reentrega" || y.Descricao == "Pendente de Entrega")
                            {
                                dataCancelamentoProtheus = util.eNulo(y.Data.ToString("yyyyMMdd"));
                            }
                            //Verificando se tem alguma timeline de 'ENTREGA' para atribuir a data de entrega
                            if (y.Descricao == "Entrega Realizada" || y.Descricao == "Reentrega Realizada")
                            {
                                dataEntregaProteus = util.eNulo(y.Data.ToString("yyyyMMdd"));
                                objOcorrencia.DataEntrega = util.eNulo(y.Data.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                            }
                        }
                        //Atribuindo a data de previsão de entrega
                        previsaoEntregaProtheus = util.eNulo(i.PrevisaoEntrega.ToString("yyyyMMdd"));
                        objOcorrencia.PrevisaoEntrega = util.eNulo(i.PrevisaoEntrega.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                        
                        objOcorrencia.StatusEntrega = util.eNulo(i.Status);
                        objOcorrencia.Obs = util.eNulo(i.UltimaOcorrencia);
                    }

                    ++processado;
                    //Validações para atualizar a tabela do Protheus
                    if (dataEntregaProteus != "")
                    {
                        ++entregue;
                        AcessoDados.AcessaDadosSQL.updateDataEntregaNf(objOcorrencia.NumNf, dataEntregaProteus);
                    }

                    if (previsaoEntregaProtheus != "")
                    {
                        AcessoDados.AcessaDadosSQL.updatePrevisaoEntregaNf(objOcorrencia.NumNf, previsaoEntregaProtheus);
                    }

                    if (dataCancelamentoProtheus != "")
                    {
                        AcessoDados.AcessaDadosSQL.updateCancelamentoEntregaNf(objOcorrencia.NumNf, dataCancelamentoProtheus);
                    }

                    DataTable dtVerUltimoStatus = AcessoDados.AcessaDadosSQL.requisitaUltimoStatus(objOcorrencia.NumNf);

                    if (dtVerUltimoStatus.Rows.Count > 0)
                    {
                        //Verificando se o status está igual, para não precisar reprocessar
                        if (dtVerUltimoStatus.Rows[0]["STATUS"].ToString() == objOcorrencia.StatusEntrega)
                        {
                            processaBD = false;
                        }
                    }

                    if (processaBD == true)
                    {
                        ++updated;
                        DataTable dtPedidos = AcessoDados.AcessaDadosSQL.requisitaPedidosProcessamentoNfs(objOcorrencia.NumNf);

                        foreach (DataRow row in dtPedidos.Rows)
                        {
                            AcessoDados.AcessaDadosSQL.insertOcorrenia(objOcorrencia, row["PEDIDO"].ToString());
                        }
                    }

                    Task.WhenAll(tasks);
                    termino = DateTime.Now;
                }

                if (total != processado)
                    erro_robo.envia_email_erro("O número de NFs processadas difere do total", null, "processaNfsAsync()");
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}