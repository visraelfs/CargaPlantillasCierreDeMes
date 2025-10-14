using CargaPlantillasCierreDeMes.Model.GestionFacturas;
using CargaPlantillasCierreDeMes.Model.ModelViews;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CargaPlantillasCierreDeMes.Clases
{
    internal class CargaPlantillaProvisiones
    {
        public string _rutaArchivo { get; set; }
        private Action<int> _reportProgress;

        public CargaPlantillaProvisiones(string rutaArchivo, Action<int> reportProgress)
        {
            _rutaArchivo = rutaArchivo;
            _reportProgress = reportProgress;
        }


        public bool CargarPlantilla()
        {
            try
            {
                _reportProgress(1);
                List<ExcelProvisiones> lregistrosExcel = ObtenerInformacionArchivo();

                var encontrarDuplicados = lregistrosExcel.GroupBy(x => new { x.LineItemId, x.Mes, x.Anio }).Where(g => g.Count() > 1).ToList();

                _reportProgress(49);

                if (encontrarDuplicados.Any())
                {
                    StringBuilder registrosDuplicados = new StringBuilder();
                    Console.WriteLine("Se encontraron duplicados:");
                    foreach (var grupo in encontrarDuplicados)
                    {

                        registrosDuplicados.AppendLine($"Producto: {grupo.Key.LineItemId}, mes: {grupo.Key.Mes}, año: {grupo.Key.Anio}");
                    }

                    throw new Exception($"Se encontraron los siguientes registros duplicados: {registrosDuplicados.ToString()}");
                }

                _reportProgress(50);

                guardarInformacionArchivoCierreDeMes(lregistrosExcel);


                return true;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        private List<ExcelProvisiones> ObtenerInformacionArchivo()
        {
            try
            {
                var workbook = new XLWorkbook(_rutaArchivo);
                var hoja = workbook.Worksheet(1);


                #region Validamos el nombre de los encabezados

                bool validarNombresEncabezados = true;

                if (hoja.Cell("A1").Value.ToString() != "LineItemId")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("B1").Value.ToString() != "IdSAE")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("C1").Value.ToString() != "NombreCuenta")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("D1").Value.ToString() != "RazonSocial")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("E1").Value.ToString() != "ExpedienteProyecto")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("F1").Value.ToString() != "ReferenciaElara")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("G1").Value.ToString() != "EstatusCoS")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("H1").Value.ToString() != "NombreProducto")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("I1").Value.ToString() != "Cantidad")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("J1").Value.ToString() != "ImporteUnitario")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("K1").Value.ToString() != "ImporteOriginal")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("L1").Value.ToString() != "DivisaCotizacion")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("M1").Value.ToString() != "ConsideradoProvision")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("N1").Value.ToString() != "FechaConsideradoProvision")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("O1").Value.ToString() != "FormaCobro")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("P1").Value.ToString() != "TipoCambio")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("Q1").Value.ToString() != "ImporteMXN")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("R1").Value.ToString() != "FechaInicio")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("S1").Value.ToString() != "Mes")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("T1").Value.ToString() != "Anio")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("U1").Value.ToString() != "SectorCliente")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("V1").Value.ToString() != "TipoIngreso")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("W1").Value.ToString() != "KAM")
                {
                    validarNombresEncabezados = false;
                }
                if (hoja.Cell("X1").Value.ToString() != "NombreProyecto")
                {
                    validarNombresEncabezados = false;
                }

                if (hoja.Cell("Y1").Value.ToString() != "FechaDeCierre")
                {
                    validarNombresEncabezados = false;
                }

                if (!validarNombresEncabezados)
                {
                    throw new Exception("Algunos de los encabzados no es correcto, valores y ordenes permitidos: LineItemId, IdSAE, NombreCuenta, RazonSocial, ExpedienteProyecto, ReferenciaElara, EstatusCoS, NombreProducto, Cantidad, ImporteUnitario, ImporteOriginal, DivisaCotizacion, ConsideradoProvision, FechaConsideradoProvision, FormaCobro, TipoCambio, ImporteMXN, FechaInicio, Mes, Anio, SectorCliente, TipoIngreso, KAM, NombreProyecto, FechaDeCierre");
                }


                #endregion

                //Empezamos a extraer la información desde la fila 2

                List<ExcelProvisiones> lregistrosExcel = new List<ExcelProvisiones>();
                ExcelProvisiones registroExcel = new ExcelProvisiones();

                int row = 2;

                foreach (var fila in hoja.RowsUsed())
                {
                    registroExcel = new ExcelProvisiones();

                    registroExcel.LineItemId = hoja.Cell($"A{row}").Value.ToString();
                    registroExcel.IdSAE = short.TryParse(hoja.Cell($"B{row}").Value.ToString(), out short idSae) ? idSae : throw new Exception($"El idSae en la fila {row} no es correcto");
                    registroExcel.NombreCuenta = hoja.Cell($"C{row}").Value.ToString();
                    registroExcel.RazonSocial = hoja.Cell($"D{row}").Value.ToString();
                    registroExcel.ExpedienteProyecto = hoja.Cell($"E{row}").Value.ToString();
                    registroExcel.ReferenciaElara = hoja.Cell($"F{row}").Value.ToString();
                    registroExcel.EstatusCoS = hoja.Cell($"G{row}").Value.ToString();
                    registroExcel.NombreProducto = hoja.Cell($"H{row}").Value.ToString();
                    registroExcel.Cantidad = short.TryParse(hoja.Cell($"I{row}").Value.ToString(), out short cantidad) ? cantidad : throw new Exception($"La cantidad en la fila {row} no es correcto");
                    registroExcel.ImporteUnitario = decimal.TryParse(hoja.Cell($"J{row}").Value.ToString(), out decimal ImporteUnitario) ? ImporteUnitario : throw new Exception($"El importe unitario en la fila {row} no es valida");
                    registroExcel.ImporteOriginal = decimal.TryParse(hoja.Cell($"K{row}").Value.ToString(), out decimal ImporteOriginal) ? ImporteOriginal : throw new Exception($"El importe original en la fila {row} no es valida");
                    registroExcel.DivisaCotizacion = hoja.Cell($"L{row}").Value.ToString();
                    registroExcel.ConsideradoProvision = hoja.Cell($"M{row}").Value.ToString();
                    registroExcel.FechaConsideradoProvision = DateTime.TryParse(hoja.Cell($"N{row}").Value.ToString(), out DateTime FechaConsideradoProvision) ? FechaConsideradoProvision : throw new Exception($"La fecha considerado provisión {row} no es valida");
                    registroExcel.FormaCobro = hoja.Cell($"O{row}").Value.ToString();
                    registroExcel.TipoCambio = decimal.TryParse(hoja.Cell($"P{row}").Value.ToString(), out decimal tipoCambio) ? tipoCambio : throw new Exception($"El tipo de cambio en la fila {row} no es correcto");
                    registroExcel.ImporteMXN = decimal.TryParse(hoja.Cell($"Q{row}").Value.ToString(), out decimal importeMXN) ? importeMXN : throw new Exception($"El Importe MXN en la fila {row} no es correcto");
                    registroExcel.FechaInicio = DateTime.TryParse(hoja.Cell($"R{row}").Value.ToString(), out DateTime fechaInicio) ? fechaInicio : throw new Exception($"La fecha inicio {row} no es valida");
                    registroExcel.Mes = short.TryParse(hoja.Cell($"S{row}").Value.ToString(), out short mes) ? mes : throw new Exception($"El mes {row} no es valido"); ;
                    registroExcel.Anio = short.TryParse(hoja.Cell($"T{row}").Value.ToString(), out short anio) ? anio : throw new Exception($"El año {row} no es valido"); ;
                    registroExcel.SectorCliente = hoja.Cell($"U{row}").Value.ToString();
                    registroExcel.TipoIngreso = hoja.Cell($"V{row}").Value.ToString();
                    registroExcel.KAM = hoja.Cell($"W{row}").Value.ToString();
                    registroExcel.NombreProyecto = hoja.Cell($"X{row}").Value.ToString();
                    registroExcel.FechaDeCierre = DateTime.TryParse(hoja.Cell($"Y{row}").Value.ToString(), out DateTime FechaDeCierre) ? FechaDeCierre : throw new Exception($"La fecha cierre en la fila {row} no es valida");

                    lregistrosExcel.Add(registroExcel);

                    row++;

                    _reportProgress(Convert.ToInt32((row * 48) / hoja.RowsUsed().Count()));

                    if (row > hoja.RowsUsed().Count())
                        break;
                }




                return lregistrosExcel;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void guardarInformacionArchivoCierreDeMes(List<ExcelProvisiones> lregistrosExcel)
        {
            try
            {
                using (Elara_GestionFacturasEntities dbContext = new Elara_GestionFacturasEntities())
                {
                    using (DbContextTransaction transaccion = dbContext.Database.BeginTransaction())
                    {
                        try
                        {
                            Provision cierreDeMesRegistro = new Provision();
                            int contadorRegistros = 0;
                            foreach (ExcelProvisiones registroExcel in lregistrosExcel)
                            {

                                Provision registroActual = dbContext.Provision.Where(x => x.LineItemId == registroExcel.LineItemId &&
                                                                                                        x.Mes == registroExcel.Mes &&
                                                                                                        x.Anio == registroExcel.Anio).FirstOrDefault();

                                contadorRegistros++;


                                if (registroActual == null)
                                {
                                    cierreDeMesRegistro = new Provision();
                                    cierreDeMesRegistro.LineItemId = registroExcel.LineItemId;
                                    cierreDeMesRegistro.IdSAE = registroExcel.IdSAE;
                                    cierreDeMesRegistro.NombreCuenta = registroExcel.NombreCuenta;
                                    cierreDeMesRegistro.RazonSocial = registroExcel.RazonSocial;
                                    cierreDeMesRegistro.ExpedienteProyecto = registroExcel.ExpedienteProyecto;
                                    cierreDeMesRegistro.ReferenciaElara = registroExcel.ReferenciaElara;
                                    cierreDeMesRegistro.EstatusCoS = registroExcel.EstatusCoS;
                                    cierreDeMesRegistro.NombreProducto = registroExcel.NombreProducto;
                                    cierreDeMesRegistro.Cantidad = registroExcel.Cantidad;
                                    cierreDeMesRegistro.ImporteUnitario = registroExcel.ImporteUnitario;
                                    cierreDeMesRegistro.ImporteOriginal = registroExcel.ImporteOriginal;
                                    cierreDeMesRegistro.DivisaCotizacion = registroExcel.DivisaCotizacion;
                                    cierreDeMesRegistro.ConsideradoProvision = registroExcel.ConsideradoProvision;
                                    cierreDeMesRegistro.FechaConsideradoProvision = registroExcel.FechaConsideradoProvision;
                                    cierreDeMesRegistro.FormaCobro = registroExcel.FormaCobro;
                                    cierreDeMesRegistro.TipoCambio = registroExcel.TipoCambio;
                                    cierreDeMesRegistro.ImporteMXN = registroExcel.ImporteMXN;
                                    cierreDeMesRegistro.FechaInicio = registroExcel.FechaInicio;
                                    cierreDeMesRegistro.Mes = registroExcel.Mes;
                                    cierreDeMesRegistro.Anio = registroExcel.Anio;
                                    cierreDeMesRegistro.SectorCliente = registroExcel.SectorCliente;
                                    cierreDeMesRegistro.TipoIngreso = registroExcel.TipoIngreso;
                                    cierreDeMesRegistro.KAM = registroExcel.KAM;
                                    cierreDeMesRegistro.NombreProyecto = registroExcel.NombreProyecto;
                                    cierreDeMesRegistro.FechaDeCierre = registroExcel.FechaDeCierre;
                                    cierreDeMesRegistro.FG = DateTime.Now;
                                    cierreDeMesRegistro.UG = "HerramientaCargaMasiva";
                                    cierreDeMesRegistro.ST = true;

                                    dbContext.Provision.Add(cierreDeMesRegistro);
                                }
                                else
                                {
                                    registroActual.ImporteUnitario = registroExcel.ImporteUnitario;
                                    registroActual.ImporteOriginal = registroExcel.ImporteOriginal;
                                    registroActual.ImporteMXN = registroExcel.ImporteMXN;
                                    registroActual.UM = "HerramientaCargaMasiva";
                                    registroActual.FM = DateTime.Now;

                                }
                                dbContext.SaveChanges();

                                //AL porcentaje obtenido se le suma 50, porque se considera que a partir del 50% es la inserción de registros
                                _reportProgress(Convert.ToInt32(((contadorRegistros * 50) / lregistrosExcel.Count()) + 50));
                            }

                            transaccion.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaccion.Rollback();
                            throw ex;
                        }
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}
