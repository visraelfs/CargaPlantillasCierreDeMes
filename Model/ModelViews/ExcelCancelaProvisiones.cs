using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CargaPlantillasCierreDeMes.Model.ModelViews
{
    internal class ExcelCancelaProvisiones
    {
        public string LineItemId { get; set; }
        public short IdSAE { get; set; }
        public string NombreCuenta { get; set; }
        public string RazonSocial { get; set; }
        public string ExpedienteProyecto { get; set; }
        public string ReferenciaElara { get; set; }
        public string EstatusCoS { get; set; }
        public string NombreProducto { get; set; }
        public short Cantidad { get; set; }
        public decimal ImporteUnitario { get; set; }
        public decimal ImporteOriginal { get; set; }
        public string DivisaCotizacion { get; set; }
        public string FormaCobro { get; set; }
        public decimal TipoCambio { get; set; }
        public decimal ImporteMXN { get; set; }
        public DateTime FechaInicio { get; set; }
        public short Mes { get; set; }
        public short Anio { get; set; }
        public string SectorCliente { get; set; }
        public string TipoIngreso { get; set; }
        public string KAM { get; set; }
        public string NoFacSae { get; set; }
        public string NombreProyecto { get; set; }
        public DateTime? FechaDeCierre { get; set; }        
    }
}
