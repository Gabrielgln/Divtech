// Root myDeserializedClass = JsonConvert.DeserializeObject<List<Root>>(myJsonResponse);
using Newtonsoft.Json;
using System.Collections.Generic;
using System;

public class NotaFiscalModel
{
    [JsonProperty("numero")]
    public string Numero { get; set; }

    [JsonProperty("origem")]
    public string Origem { get; set; }

    [JsonProperty("destino")]
    public string Destino { get; set; }

    [JsonProperty("emissao")]
    public DateTime Emissao { get; set; }

    [JsonProperty("remetente")]
    public string Remetente { get; set; }

    [JsonProperty("destinatario")]
    public string Destinatario { get; set; }

    [JsonProperty("tipoFrete")]
    public string TipoFrete { get; set; }

    [JsonProperty("volumes")]
    public int Volumes { get; set; }

    [JsonProperty("valorMercantil")]
    public double ValorMercantil { get; set; }

    [JsonProperty("peso")]
    public double Peso { get; set; }

    [JsonProperty("totalFrete")]
    public double TotalFrete { get; set; }

    [JsonProperty("previsaoEntrega")]
    public DateTime PrevisaoEntrega { get; set; }

    [JsonProperty("dataEntrega")]
    public DateTime DataEntrega { get; set; }

    [JsonProperty("status")]
    public string Status { get; set; }

    [JsonProperty("cidade")]
    public string Cidade { get; set; }

    [JsonProperty("uf")]
    public string Uf { get; set; }

    [JsonProperty("cidadeColeta")]
    public string CidadeColeta { get; set; }

    [JsonProperty("ufColeta")]
    public string UfColeta { get; set; }

    [JsonProperty("dataOcorrencia")]
    public DateTime DataOcorrencia { get; set; }

    [JsonProperty("ultimaOcorrencia")]
    public string UltimaOcorrencia { get; set; }

    [JsonProperty("notasFiscais")]
    public List<NotasFiscais> NotasFiscais { get; set; }

    [JsonProperty("timeline")]
    public List<Timeline> Timeline { get; set; }

    [JsonProperty("ocorrencias")]
    public List<object> Ocorrencias { get; set; }
}

public class NotasFiscais
{
    [JsonProperty("serie")]
    public int Serie { get; set; }

    [JsonProperty("numero")]
    public int Numero { get; set; }

    [JsonProperty("emissao")]
    public DateTime Emissao { get; set; }
}

public class Timeline
{
    [JsonProperty("descricao")]
    public string Descricao { get; set; }

    [JsonProperty("data")]
    public DateTime Data { get; set; }
}

