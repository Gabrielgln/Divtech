using ProcessamentoRoboClasses.ApiSistemaComercial2.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessamentoRoboClasses.ApiSistemaComercial2.Interfaces
{
    internal interface INotaFiscal
    {
        Task<List<NotaFiscalModel>> BuscarNotaFiscalPorNumero(int numero);


    }
}