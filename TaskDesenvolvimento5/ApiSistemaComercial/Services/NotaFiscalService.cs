using Newtonsoft.Json;
using ProcessamentoRoboClasses.ApiSistemaComercial2.Models;
using ProcessamentoRoboClasses.ApiSistemaComercial2.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using CsQuery.ExtensionMethods.Internal;

namespace ProcessamentoRoboClasses.ApiSistemaComercial2.Services
{
    internal class NotaFiscalServices : INotaFiscal
    {
        public static string Token { get; set; }
        public static DateTime ExpirationDate { get; set; }

        public static void getTokenByLogin()
        {
            try
            {
                string urlApiGeraToken = "https://api.tessindustria.com.br/Login/Authenticate";

                using (var client = new HttpClient())
                {
                    string usuario = "wjgWV2W.fOEABHD5aeJL95pkG3Q2tg8kDgeNbPxkKSPeE5WA0orwthQZ16ViIFqhqlskB6IRtdw" +
                        "JHYkT0bMmlLFie7TEyWqZqEqAXLMtSLvT4nLr4t5PWrwNK6SsaAYX1dr0HZBL54vs6lXMBHNL91KMFfI2dZ86ii8" +
                        "s1CdwttcQE6ncef7wRKkEygj6kConlcU240KCeYsgLFNKRQYvjOxBi8t6OSwWQs8uTvO4QHkJqhJfZApwOGUUS35" +
                        "zUHVt";
                    string senha = "ywJUjYJRG8Pr8N8qgS";

                    var Login = new
                    {
                        apiKey = usuario,
                        password = senha
                    };

                    string JsonObjeto = JsonConvert.SerializeObject(Login);
                    var content = new StringContent(JsonObjeto, Encoding.UTF8, "application/json");

                    var response = client.PostAsync(urlApiGeraToken, content).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var tokenJsonString = response.Content.ReadAsStringAsync();
                        tokenJsonString.Wait();
                        var authenticate = JsonConvert.DeserializeObject<Authenticate>(tokenJsonString.Result);
                        Token = authenticate.AccessToken.ToString();
                        var ExpirationString = authenticate.Expiration.ToString();
                        ExpirationDate = DateTime.Parse(ExpirationString);
                        
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public bool TokenInvalid()
        {
            bool tokenInvalid = false;
            if (ExpirationDate.ToString() == "01/01/0001 00:00:00")
            {
                tokenInvalid = true;
            }
            else
            {
                DateTime dataHoraDaAplicacao = DateTime.Now;
                TimeSpan timeSpan = ExpirationDate - dataHoraDaAplicacao;
                if(timeSpan.TotalMinutes <= 0)
                {
                    tokenInvalid = true;
                }
            }
            return tokenInvalid;
        }

        public async Task<List<NotaFiscalModel>> BuscarNotaFiscalPorNumero(int numNotaFiscal)
        {
            try
            {
                if (TokenInvalid())
                {
                    getTokenByLogin(); //Gerar token
                }
                
                if (!string.IsNullOrEmpty(Token)) //Verifica se o Token não é nulo o vázio
                {
                    //Pega a url da API, que realiza essa função "RastreamentoPedidoPorNotaFiscal"
                    string urlApiSistemaComercial = $"https://api.tessindustria.com.br/Comercial/Transportadoras/RastreamentoPedidoPorNotaFiscal/{numNotaFiscal}";
                    //Criando um classe "System.Net.Http" para enviar e receber solicitações HTTP
                    using (var httpClient = new HttpClient())
                    {
                        //Limpando o Header, para não acumular "Tokens"
                        httpClient.DefaultRequestHeaders.Clear();
                        //Solicitando autorização do Header com o "Token"
                        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", Token);
                        //Pegando as informações da função "RastreamentoPedidoPorNotaFiscal" com o Http.get
                        var request = new HttpRequestMessage(HttpMethod.Get, urlApiSistemaComercial);
                        //aguardando a resposta do servidor
                        var response = await httpClient.SendAsync(request);
                        //pegando essa resposta em formato de string do tipo JSON
                        var notaFiscalJsonString = await response.Content.ReadAsStringAsync();
                        //transformar o json de string para um objeto NotaFiscal do tipo List
                        var notaFiscalModelList = JsonConvert.DeserializeObject<List<NotaFiscalModel>>(notaFiscalJsonString).ToList();

                        if (notaFiscalModelList != null)
                        {
                            return notaFiscalModelList;
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}