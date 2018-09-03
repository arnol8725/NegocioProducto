using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Elektra.Negocio.Entidades.Producto;
using Elektra.Negocio.Entidades.SistemasInformacion.Productos;

using System.Globalization;
using System.Data;
using EntidadMileniaPostVenta;
using System.Collections;
using Elektra.Negocio.EntregaDomicilio;
using Elektra.Negocio.Entidades.Ventas;
using Microsoft.ApplicationBlocks.ConfigurationManagement;
using Elektra.Negocio.Entidades.Milenia;



namespace Elekta.Negocio.SistnformacionEpos
{
    public class ManejadorSistemasInformacion
    {


        public DetalleProducto consultaProducto(int tipoBusqueda, string consulta)
        {
            DetalleProducto respuesta = new DetalleProducto();
            respuesta.detBusquesa = new List<DetalleBusqueda>();
            try {

                switch (tipoBusqueda) { 
                    case 1 : //Codigo
                        respuesta.detBusquesa = convertirDetalleBusqueda(BuscarProducto(consulta));
                        break;
                    case 2: //Descripcion
                        respuesta.detBusquesa = convertirDetalleBusqueda(BuscarProductoPalabra(consulta));
                        break;
                    case 3: //Marca
                        respuesta.detBusquesa = convertirDetalleBusqueda(BuscarProductoMarca(consulta));
                        break;

                }
            
            }
            catch (Exception e) {
                respuesta.error = true;
                respuesta.mensaje= "Error al realizar la busqueda del producto";
                respuesta.mensajeTecnico = "Error al realizar la busqueda del producto: "+e.Message;
            }

            return respuesta;                 
        }

        public List<DetalleBusqueda> convertirDetalleBusqueda(DataSet detalle) {
            List<DetalleBusqueda> respuesta = new List<DetalleBusqueda>();
            if (detalle.Tables != null || detalle.Tables[0].Rows.Count > 0) {

                foreach (DataRow productos in detalle.Tables[0].Rows) {
                    DetalleBusqueda det = new DetalleBusqueda();
                    det.codigo = Double.Parse(productos["fiProdId"].ToString());
                    det.descripcion = productos["fcProdDesc"].ToString();
                    det.marca = productos["fcMarca"].ToString();
                    det.precio = productos["fnProdPrecio"].ToString();
                    det.precio = String.Format("{0:00}", det.precio);

                    respuesta.Add(det);
                }
                                
            }


            return respuesta;

        }

        public DataSet EjecutarCriteriosPostBusqueda(DataSet productos)
        {
            DataRow[] renglones = productos.Tables[0].Select("fcProdDesc LIKE '%plan inteligente%'");
            if (renglones.Length >= 1)
            {
                DataTable tablaNueva = productos.Tables[0].Clone();
                DataRow[] renglonesFiltrados = productos.Tables[0].Select("fcProdDesc NOT LIKE '%plan inteligente%'");
                foreach (DataRow dr in renglonesFiltrados)
                {
                    tablaNueva.ImportRow(dr);
                }
                DataSet datosProductos = new DataSet();
                datosProductos.Tables.Add(tablaNueva);
                productos = datosProductos;
            }

            return productos;
        }
        public DataSet BuscarProducto(string datoBusqueda)
        {
            EntProducto producto = new EntProducto();
            producto.FiltraParaLaborVenta = false;
          
            try
            {
                DataSet ds = producto.ObtenerDataSetProductoPorSku(Convert.ToInt32(datoBusqueda));
                if (ds != null)
                {
                    DataSet resultadoFinal = this.EjecutarCriteriosPostBusqueda(ds);
                    if (resultadoFinal != null)
                        return resultadoFinal;
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        public DataSet BuscarProductoPalabra(string datoBusqueda)
        {
            EntProducto producto = new EntProducto();
            producto.FiltraParaLaborVenta = true;
            DataSet ds = producto.ObtenerProductoPorPalabraClave(datoBusqueda);
            return this.EjecutarCriteriosPostBusqueda(ds);
        }

        public DataSet BuscarProductoMarca(string datoBusqueda)
        {
            EntProducto producto = new EntProducto();
            producto.FiltraParaLaborVenta = true;
            DataSet productos = null;

            productos = producto.ObtenerProductoPorMarca(datoBusqueda);

            return this.EjecutarCriteriosPostBusqueda(productos);
        }


        #region Metodo para la carga del detalle del producto
        public DetalleProductoxSKU detalleProducto(int sku, string empleado)
        {
            Respuesta resp = new Respuesta();
            DetalleProductoxSKU det = new DetalleProductoxSKU();
            det.detExistInv = new List<ExistenciaInventario>();
            det.detMov = new List<DetalleMovimiento>();
            det.detVtaGeneral = new List<VentaGeneral>();
            det.detInvTipo = new List<ExistenciaInventarioxTipo>();
            try {
                var productoSistemasInformacion = new EntSistemasInformacionProductos(sku);
                productoSistemasInformacion.Empleado = empleado;

                det.codigo = productoSistemasInformacion.Producto.Id.ToString();
                det.descripcion = productoSistemasInformacion.Producto.Descripcion.ToString();
                det.negocio = productoSistemasInformacion.Producto.Negocio.Descripcion.ToString();
                det.depto = productoSistemasInformacion.DescripcionDepartamento.ToString();
                det.subDepto= productoSistemasInformacion.DescripcionSubDepartamento.ToString();
                det.clase = productoSistemasInformacion.DescripcionClase.ToString();
                det.subClase= productoSistemasInformacion.DescripcionSubClase.ToString();

                det.tipoProducto = productoSistemasInformacion.TipoProducto.ToString();
                det.estado = productoSistemasInformacion.Producto.EsActivo ? "Activo" : "Inactivo";
                det.fechaAlta = det.estado == "Activo" || productoSistemasInformacion.FechaBaja.Year == 1900 ? productoSistemasInformacion.FechaAlta.ToShortDateString() : productoSistemasInformacion.FechaBaja.ToShortDateString();
                det.precioContado = productoSistemasInformacion.Producto.PrecioChaz.ToString();
                det.descMaximo = productoSistemasInformacion.Producto.PorcentajeDescuentoContado == 0 ? "0%" : productoSistemasInformacion.Producto.PorcentajeDescuentoContado.ToString("###.##") + "%";

                det.ultCambioPrecio = productoSistemasInformacion.FechaUltimoCambioPrecio.Year > 1900 ? productoSistemasInformacion.FechaUltimoCambioPrecio.ToShortDateString() : "";
                det.tiempoGarantia = productoSistemasInformacion.TiempoGarantia.ToString();
                det.diasReparacion = productoSistemasInformacion.DiasReparacion.ToString();
                det.paquete =  productoSistemasInformacion.Paquete.ToString();
                det.manejaNoSerie = productoSistemasInformacion.Producto.ClasificacionJda.EsRequeridoNumeroSerie ? "SI" : "NO";
                det.codigoBarras = productoSistemasInformacion.CodigoBarras.ToString();
                det.tipoBloqueo = productoSistemasInformacion.Producto.BloqueoJda.ToString();

                var fechaHoy = DateTime.Today;
               
			    var fechaInicialMovimientosInventario = new DateTime(fechaHoy.Year, 1, 01, 0, 0, 0, 0);
			    var fechaFinalMovimientosInventario = new DateTime(fechaHoy.Year, 12, 31, 23, 59, 59, 999);
                productoSistemasInformacion.FechaInicialMovimientosInventario = fechaInicialMovimientosInventario;			
				productoSistemasInformacion.FechaFinalMovimientosInventario = fechaFinalMovimientosInventario;

                var detalle = productoSistemasInformacion.DetalleMovimientosInventario;
                List<DetalleMovimiento> detMov= new List<DetalleMovimiento> ();


                if (detalle.Tables[0] != null) {
                    foreach (DataRow deta in detalle.Tables[0].Rows)
                    {
                        DetalleMovimiento mov = new DetalleMovimiento();
                        mov.NoTransac=Int32.Parse(deta["fiNoTransac"].ToString());   
                         mov.UbiId         =deta["fiUbiId"].ToString() ;
                         mov.DMInvCant     =deta["fiDMInvCant"].ToString() ;
                         mov.TipoOp        =deta["fiTipoOp"].ToString() ;
                         mov.TopDesc       =deta["fcTopDesc"].ToString() ;
                         mov.MInvRef       =deta["fcMInvRef"].ToString() ;
                         mov.ExistIni      =deta["fiExistIni"].ToString() ;
                         mov.ExistFin      =deta["fiExistFin"].ToString() ;
                         mov.MInvEntSal    =deta["fcMInvEntSal"].ToString() ;
                         mov.EmpIDRed      =deta["fcEmpIDRed"].ToString() ;
                         mov.WS            =deta["fcWS"].ToString() ;
                         mov.UbiDesc       =deta["fcUbiDesc"].ToString() ;
                         mov.fecha         =deta["fecha"].ToString() ;
                         det.detMov.Add(mov);
                    }
                }


                #region Detalle de Inventario
                var detInv= productoSistemasInformacion.ExistenciasInventario;
                    if (detInv.Tables[0] != null) {
                        foreach (DataRow datos in detInv.Tables[0].Rows)
                        {
                            ExistenciaInventario Inv = new ExistenciaInventario();
                            Inv.ubicacion = Int32.Parse(datos["fiUbiId"].ToString()); 
                            Inv.descripcion = datos["fcUbiDesc"].ToString();
                            Inv.existencia = Int32.Parse(datos["fiInvExist"].ToString());
                            det.detExistInv.Add(Inv);
                        }                    
                        
                    }

                    if (detInv.Tables[1] != null)
                    {
                        foreach (DataRow datos in detInv.Tables[1].Rows)
                        {
                            ExistenciaInventarioxTipo Inv = new ExistenciaInventarioxTipo();
                            Inv.Contadosinsurtir = Int32.Parse(datos["ExistenciaContadoSinSurtir"].ToString());
                            Inv.Apartadocon90pagado = Int32.Parse(datos["ExistenciaApartadoAlNoventaPorCiento"].ToString());
                            Inv.Apartadoconmenosde90pagado = Int32.Parse(datos["ExistenciaApartadoMenorAlNoventaPorCiento"].ToString());
                            Inv.Creditocon90deenganchepagado = Int32.Parse(datos["ExistenciaEngancheAlNoventaPorCiento"].ToString());


                            det.detInvTipo.Add(Inv);
                        }

                    }
                #endregion

                    #region Detalle general
                    var detGeneral = productoSistemasInformacion.VentasGeneradas;
                    if (detGeneral.Tables[0] != null)
                    {
                        foreach (DataRow datos in detGeneral.Tables[0].Rows)
                        {
                            VentaGeneral Inv = new VentaGeneral();
                            Inv.marcadas = Int32.Parse(datos["MarcadasMesActual"].ToString());
                            Inv.surtidas = Int32.Parse(datos["SurtidasMesActual"].ToString());
                            Inv.canceladasADE = Int32.Parse(datos["CAEMesActual"].ToString());
                            Inv.canceladasDDE = Int32.Parse(datos["CDEMesActual"].ToString());
                            det.detVtaGeneral.Add(Inv);
                            Inv = new VentaGeneral();
                            Inv.marcadas = Int32.Parse(datos["MarcadasMesAnterior"].ToString());
                            Inv.surtidas = Int32.Parse(datos["SurtidasMesAnterior"].ToString());
                            Inv.canceladasADE = Int32.Parse(datos["CAEMesAnterior"].ToString());
                            Inv.canceladasDDE = Int32.Parse(datos["CDEMesAnterior"].ToString());
                            det.detVtaGeneral.Add(Inv);
                        }

                    }
                    #endregion



            }
            catch (Exception e)
            {
                resp.error = true;
                resp.mensaje = "Error para obtener detalle de producto";
                resp.mensajeTecnico  = "Error para obtener detalle de producto";

            }
            return det;
            
        }

        #endregion




        #region Metodoss milanesas

        private StringBuilder XmlMileniaSingle(string fileXls,
        string poliza,
        string nombreCliente,
        string direccionCliente,
        string telefonoCliente,
        string garantiaExtendida,
        DateTime inicioVigencia,
        DateTime finVigencia,
        string articuloDescripcion,
        string esAplicaImpresionSerie,
        string numeroSerie,
        DateTime nuevasClausulas)
            {
                StringBuilder sbContenido = new StringBuilder();
                sbContenido = new StringBuilder();
                sbContenido.Append("<?xml version='1.0' encoding='utf-8'?>");
                sbContenido.Append("<?xml-stylesheet type='text/xsl' href='" + fileXls + "'?>");
                sbContenido.Append("<Datos>");
                sbContenido.Append("<DatosImpresion>");
                sbContenido.Append("<Tipo>cta</Tipo>");
                sbContenido.Append("</DatosImpresion>");
                sbContenido.Append("<GarantiaExtendida>");
                sbContenido.Append("<DatosGarantia>");

                if (inicioVigencia >= nuevasClausulas)
                    sbContenido.Append("<NvaClausulas>1</NvaClausulas>");
                else
                    sbContenido.Append("<NvaClausulas>0</NvaClausulas>");

                sbContenido.Append("<NoPoliza>" + poliza + "</NoPoliza>");
                sbContenido.Append("<NombreCliente>" + nombreCliente + "</NombreCliente>");
                sbContenido.Append("<DomicilioCliente>" + direccionCliente + "</DomicilioCliente>");
                sbContenido.Append("<TelefonoCliente>" + telefonoCliente + "</TelefonoCliente>");
                sbContenido.Append("<DuracionPoliza>" + garantiaExtendida + "</DuracionPoliza>");
                sbContenido.Append("<FechaInicio>" + inicioVigencia.Day.ToString() + " de " + inicioVigencia.ToString("MMMM", CultureInfo.CurrentCulture) + " de " + inicioVigencia.Year.ToString() + "</FechaInicio>");
                sbContenido.Append("<FechaFin>" + finVigencia.Day.ToString() + " de " + finVigencia.ToString("MMMM", CultureInfo.CurrentCulture) + " de " + finVigencia.Year.ToString() + "</FechaFin>");
                sbContenido.Append("<NombreCompania>Elektra del Milenio SA de CV</NombreCompania>");
                sbContenido.Append("<TipoTienda></TipoTienda>");
                sbContenido.Append("</DatosGarantia>");
                sbContenido.Append("<Productos>");
                sbContenido.Append("<Descripcion>" + articuloDescripcion + "</Descripcion>");
                sbContenido.Append("<EsImprimeSerie>" + esAplicaImpresionSerie.ToString() + "</EsImprimeSerie>");
                sbContenido.Append("<NumSerie>" + numeroSerie + "</NumSerie>");
                //sbContenido.Append("<Precio>" + viPrecio + "</Precio>");
                sbContenido.Append("</Productos>");
                sbContenido.Append("</GarantiaExtendida>");
                sbContenido.Append("</Datos>");

                return sbContenido;

        }//

        public bool ValidaImpresionMilenia()
        {
            bool esImpresionMilenia = false;
            EntMileniaPosVenta catalogoGenerico = new EntMileniaPosVenta();

            DataSet catalogo = catalogoGenerico.ObtenerCatalogo(57);

            if (catalogo.Tables.Count > 0)
            {
                foreach (DataRow Catalogos in catalogo.Tables[0].Rows)
                {
                    if (Catalogos["fiItemId"].ToString().Trim().Equals("17") && Convert.ToBoolean(Catalogos["flStatus"]))
                    {
                        esImpresionMilenia = true;
                        break;
                    }
                }

            }

            return esImpresionMilenia;
        }
        public bool ValidarExisteMilenia(int Pedido)
        {
            bool vbContieneMil = false;
            EntMileniaPosVenta Milenia = new EntMileniaPosVenta();
            try
            {
                DataSet esmilenia = Milenia.PedidoContieneMilenia(Pedido);
                if (esmilenia != null && esmilenia.Tables.Count > 0 && esmilenia.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToBoolean(esmilenia.Tables[0].Rows[0]["EsMilenia"]))
                    {
                        vbContieneMil = true;
                    }
                }
                return vbContieneMil;
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine("Error al verificar si es el pedido contiene Milenia ", "LOG");
               
                return vbContieneMil;
            }
        }
       /* private string[] FiltrarPolizaReimpresion(string cadena)
        {
            string[] arrayRePrintPolizaMilenia = null;
            arrayRePrintPolizaMilenia = this.polizaMilenia.Split('|');
            return arrayRePrintPolizaMilenia;
        }*/

        private bool CompararPolizaMilenia(string poliza, string[] Polizas)
        {
            bool isElement = false;
            foreach (string item in Polizas)
            {
                if (item.Equals(poliza))
                {
                    isElement = true;
                    break;
                }
            }
            return isElement;
        }

        private StringBuilder XmlMileniaDouble(string fileXls,
        string[] poliza,
        string[] nombreCliente,
        string[] direccionCliente,
        string[] telefonoCliente,
        string[] garantiaExtendida,
        DateTime[] inicioVigencia,
        DateTime[] finVigencia,
        string[] articuloDescripcion,
        string[] esAplicaImpresionSerie,
        string[] numeroSerie,
        DateTime nuevasClausulas,
        EntMileniaPosVenta EntMileniaR
        )
        {
            StringBuilder sbContenido = new StringBuilder();
            sbContenido = new StringBuilder();
            sbContenido.Append("<?xml version='1.0' encoding='utf-8' ?>");

            /*if (EntMileniaR.esMileRemplazo(impresionSurtimiento.Venta.IdPedido))
                sbContenido.Append("<?xml-stylesheet type='text/xsl' href='" + fileXls + "' ?>");
            else
                sbContenido.Append("<?xml-stylesheet type='text/xsl' href='/ElektraFront/Xsl/GarantiaExtendida.xslt' ?>");*/
            sbContenido.Append("<?xml-stylesheet type='text/xsl' href='" + fileXls + "' ?>");

            sbContenido.Append("<Datos>");
            sbContenido.Append("<DatosImpresion>");
            sbContenido.Append("<Tipo>cta</Tipo>");
            sbContenido.Append("</DatosImpresion>");
            sbContenido.Append("<GarantiaExtendida>");
            sbContenido.Append("<DatosGarantia>");

            //if (inicioVigencia >= nuevasClausulas)
            //    contenido.Append("<NvaClausulas>1</NvaClausulas>");
            //else
            //    contenido.Append("<NvaClausulas>0</NvaClausulas>");

            sbContenido.Append("<NoPoliza>" + poliza[0] + "</NoPoliza>");
            sbContenido.Append("<NombreCliente>" + nombreCliente[0] + "</NombreCliente>");
            sbContenido.Append("<DomicilioCliente>" + direccionCliente[0] + "</DomicilioCliente>");
            sbContenido.Append("<TelefonoCliente>" + telefonoCliente[0] + "</TelefonoCliente>");

            sbContenido.Append("<DuracionPoliza1>" + garantiaExtendida[0] + "</DuracionPoliza1>");
            sbContenido.Append("<FechaInicio1>" + inicioVigencia[0].Day.ToString() + " de " + inicioVigencia[0].ToString("MMMM", CultureInfo.CurrentCulture) + " de " + inicioVigencia[0].Year.ToString() + "</FechaInicio1>");
            sbContenido.Append("<FechaFin1>" + finVigencia[0].Day.ToString() + " de " + finVigencia[0].ToString("MMMM", CultureInfo.CurrentCulture) + " de " + finVigencia[0].Year.ToString() + "</FechaFin1>");

            sbContenido.Append("<DuracionPoliza2>" + garantiaExtendida[1] + "</DuracionPoliza2>");
            sbContenido.Append("<FechaInicio2>" + inicioVigencia[1].Day.ToString() + " de " + inicioVigencia[1].ToString("MMMM", CultureInfo.CurrentCulture) + " de " + inicioVigencia[1].Year.ToString() + "</FechaInicio2>");
            sbContenido.Append("<FechaFin2>" + finVigencia[1].Day.ToString() + " de " + finVigencia[1].ToString("MMMM", CultureInfo.CurrentCulture) + " de " + finVigencia[1].Year.ToString() + "</FechaFin2>");

            sbContenido.Append("<NombreCompania>Elektra del Milenio SA de CV</NombreCompania>");
            sbContenido.Append("<TipoTienda></TipoTienda>");
            sbContenido.Append("</DatosGarantia>");
            sbContenido.Append("<Productos>");

            //if (!esReimpMilPostVenta)
            //    contenido.Append("<Descripcion>" + detalleVenta.ProductoServicio.Descripcion + "</Descripcion>");
            //else
            //    contenido.Append("<Descripcion>" + garantiaExtendida["FcProdDesc"].ToString() + "</Descripcion>");

            sbContenido.Append("<Descripcion1>" + articuloDescripcion[0] + "</Descripcion1>");
            sbContenido.Append("<EsImprimeSerie1>" + esAplicaImpresionSerie[0].ToString() + "</EsImprimeSerie1>");
            sbContenido.Append("<NumSerie1>" + numeroSerie[0] + "</NumSerie1>");

            sbContenido.Append("<Descripcion2>" + articuloDescripcion[1] + "</Descripcion2>");
            sbContenido.Append("<EsImprimeSerie2>" + esAplicaImpresionSerie[1].ToString() + "</EsImprimeSerie2>");
            sbContenido.Append("<NumSerie2>" + numeroSerie[1] + "</NumSerie2>");

            sbContenido.Append("</Productos>");
            sbContenido.Append("</GarantiaExtendida>");
            sbContenido.Append("</Datos>");

            return sbContenido;

        }//

        private bool esReimpresion = false;
        private int contadorMilenia=0;

        

        public ArrayList ImprimeGarantiaLigadaXML(DataSet detalleMilenia, string ip)
        {
            ArrayList retorno = new ArrayList();
            bool esReimpMilPostVenta = true;
            double SKUReimprimir = 0;
            double SKUReImpresoMilPosTVenta = 0;
            EntMileniaPosVenta EntMileniaR = new EntMileniaPosVenta();
            //double skuPromoMilenia = Convert.ToDouble(((Hashtable)ConfigurationManager.Read("MileniaConfiguracion"))["SKUInstalacionMiniSplit"]);
            string[] Polizas = null;
            int viEstilo = 1;
            try
            {
                StringBuilder contenido;
                string nombreCliente = "";
                string telefonoCliente = "";
                string direccionCliente = "";
                int esAplicaImpresionSerie = 1;
                DataTable dtDatosMilPV = new DataTable();
                bool esMileniaPostventaEmpleado = false;
                string duracionGExt = "";
                DateTime iniVig;
                DateTime finVig;
                string descripProdcuto = "";
                bool esAntiguoMil = false;
                string poliza = "";
                int viPedido = (int)detalleMilenia.Tables[0].Rows[0]["fiNopedido"];

                ManejadorEntregaDomicilio manejadorEAD = new ManejadorEntregaDomicilio();
                if (manejadorEAD.ValidaTiendaConImpresionMilenia())
                    esAplicaImpresionSerie = 0;//Verifica si es milenia entrega a domicilio 


                EntMileniaPosVenta entmilPV = new EntMileniaPosVenta();
                DataSet dsDatosVAE = entmilPV.DatosEmpleadoPostVentaMilenia(viPedido);

                if (dsDatosVAE.Tables.Count > 0 && dsDatosVAE.Tables[0].Rows.Count > 0)
                {
                    esMileniaPostventaEmpleado = true;
                    nombreCliente = dsDatosVAE.Tables[0].Rows[0]["fcEmpNombre"].ToString().Trim();
                    telefonoCliente = dsDatosVAE.Tables[0].Rows[0]["fcTelefono"].ToString().Trim();
                    direccionCliente = dsDatosVAE.Tables[0].Rows[0]["fcCalle"].ToString().Trim() + ", " + dsDatosVAE.Tables[0].Rows[0]["fcColonia"].ToString().Trim() + ", " + dsDatosVAE.Tables[0].Rows[0]["fcEstado"].ToString().Trim();
                    dtDatosMilPV = EntMileniaR.ObtieneTiposMileniasPV(viPedido).Tables[0];
                }
                else
                {

                    nombreCliente = detalleMilenia.Tables[0].Rows[0]["fcCteNombre"].ToString().Trim() + " " +
                        detalleMilenia.Tables[0].Rows[0]["fcCteApaterno"].ToString().Trim() + " " +
                        detalleMilenia.Tables[0].Rows[0]["fcCteAMaterno"].ToString().Trim();


                    if (detalleMilenia.Tables[0].Rows[0]["fcCteTel"].ToString().Trim() != null)
                    {
                        telefonoCliente = detalleMilenia.Tables[0].Rows[0]["fcCteTel"].ToString().Trim();
                    }
                    if (detalleMilenia.Tables[0].Rows[0]["fcCteDirCalle"].ToString().Trim() != null)
                    {
                        direccionCliente = detalleMilenia.Tables[0].Rows[0]["fcCteDirCalle"].ToString().Trim() + " " +
                                           detalleMilenia.Tables[0].Rows[0]["fcCteColonia"].ToString().Trim();
                    }
                }

                EntRelacionVentaConGarantia garantia = new EntRelacionVentaConGarantia();
                DataSet dsDatosGarantias = detalleMilenia;

                ///si es un pedido de solo milenias Obtenemos la sku de la poliza a reimprimir
               /* if (esReimpMilPostVenta)
                {
                    if ((dsDatosGarantias.Tables.Count > 0) && (dsDatosGarantias.Tables[0].Rows.Count > 0))
                    {
                        foreach (DataRow garantiaExtendida in dsDatosGarantias.Tables[0].Rows)
                        {
                            if (this.polizaMilenia == garantiaExtendida["fiPoliza"].ToString())
                            {
                                SKUReimprimir = Double.Parse(garantiaExtendida["fiProdId"].ToString());
                            }
                        }
                    }
                }*/

                //Si es reimpresion, se almacenan el folio de la poliza(s) seleccionada
                /*if (this.esReimpresion)
                {
                    Polizas = FiltrarPolizaReimpresion(this.polizaMilenia);
                }*/

              

             
               

                foreach (DataTable dtMilenia in detalleMilenia.Tables)
                {


                    /*Validacion de cantida de sku*/
                    

                    foreach (DataRow garantiaExtendida in dtMilenia.Rows)
                    {


                        if (garantiaExtendida["fdFecIniGext"] == System.DBNull.Value || garantiaExtendida["fdFecIniGext"] == System.DBNull.Value)
                        {
                            continue;
                        }
                        if (garantiaExtendida["fdFecIniGext"].ToString() == garantiaExtendida["DuracionGEXT"].ToString())
                        {
                            continue;
                        }

                        poliza = garantiaExtendida["fiPoliza"].ToString();

                        //Validación que reimprime solo la poliza seleccionada, omite las no seleccionadas. 
                        if (this.esReimpresion)
                        {
                            if (!poliza.Equals(String.Empty))
                            {
                                if (!this.CompararPolizaMilenia(poliza, Polizas))
                                {
                                    continue;
                                }
                            }
                        }

                        if (this.esReimpresion)
                        {

                            /**** Si recibe un número de poliza es de reimpresión ****/
                           /* if (!esReimpMilPostVenta)
                            {
                                if (!this.polizaMilenia.Equals(String.Empty))
                                {
                                    if (!this.polizaMilenia.Equals(poliza))
                                    { continue; }
                                }
                            }*/
                        }

                        string numeroSerie = "";

                        //Validacion para saber si tiene un numero de serie o se imprime como omitido

                        //Concatena el valor del precio del producto
                        numeroSerie = garantiaExtendida["fiNoSerie"].ToString(); //+ "\n" +" Precio: $" + garantiaExtendida["viPrecio"].ToString().Trim() ; 


                        //Empieza la identificacion de tipos de polizas Milenia		
                        EntMileniaPlus EntVerificagarantia = new EntMileniaPlus();

                        ///EntRelacionVentaConGarantia EntVerificagarantia = new EntRelacionVentaConGarantia();
                        String tipoMileniaImpr = string.Empty;
                        if (esMileniaPostventaEmpleado)
                        {
                            DataRow[] drAux = dtDatosMilPV.Select("fiprodID = " + garantiaExtendida["FIPRODID"].ToString());

                            if (drAux.Length > 0)
                                tipoMileniaImpr = drAux[0]["FCTIPO"].ToString().Trim();
                            else
                                tipoMileniaImpr = "SIN TIPO";
                        }
                        else
                        {
                            tipoMileniaImpr = EntVerificagarantia.obtenerTipoMileniaImprimir(Double.Parse(garantiaExtendida["fiProdId"].ToString()), viPedido);
                        }

                        //Se valida que el tipo de milenia no venga vacio 
                        if (tipoMileniaImpr.Equals(""))
                            continue;

                        //Fecha a partir de la cual se imprimen nuevas polizas
                        DateTime nuevasClausulas = Convert.ToDateTime("04/09/2009");

                        // Se invoca al generador de XML de impresion
                        duracionGExt = garantiaExtendida["Duracion"].ToString();
                        iniVig = Convert.ToDateTime(garantiaExtendida["fdFecIniGProv"].ToString().Trim());
                        //viPrecio = garantiaExtendida["viPrecio"].ToString();
                        finVig = Convert.ToDateTime(garantiaExtendida["DuracionGEXT"].ToString().Trim());
                        descripProdcuto = garantiaExtendida["fcProdDesc"].ToString().Trim();
                        esAntiguoMil = !Convert.ToBoolean(((Hashtable)ConfigurationManager.Read("MileniaConfiguracion"))["NuevosCertificados"]);
                        /*Validacion para saber que tipo de poliza manda*/

                        DateTime FechaImpresionNuevaMil = Convert.ToDateTime(((Hashtable)ConfigurationManager.Read("MileniaConfiguracion"))["FechaImprimeNuevaMil"]);

                        if (Convert.ToBoolean(((Hashtable)ConfigurationManager.Read("MileniaConfiguracion"))["EsMileniaAsterisco"]))
                        {
                            //Valida si la la fecha es mayor a la fecha que se enceuntra en Configuracionmanager, si es mayor, significa que es una poliza nueva, si 1no , esta 
                            //activa la configuracion para que imprima las poliza nuevas con el asterisco pero la poliza sigue siendo viejita, antes de la fecha estipulada
                            if (DateTime.Now <= FechaImpresionNuevaMil)
                            {
                                //Imprime Poliza Con asterisco
                                viEstilo = 3;
                            }
                            else
                            {

                                viEstilo = 1;
                            }

                        }
                        else if (DateTime.Now >= FechaImpresionNuevaMil)
                        {
                            //Imprime Poliza Nueva sin asterisco
                            viEstilo = 2;
                        }

                        string rutaXSLT;

                        
                        
                            contenido = this.GetXmlPoliza(viEstilo, tipoMileniaImpr, poliza, nombreCliente, direccionCliente, telefonoCliente, duracionGExt, iniVig, finVig, descripProdcuto, esAplicaImpresionSerie.ToString(), numeroSerie, nuevasClausulas, esAntiguoMil, ip, out rutaXSLT);
                       
                        
                        

                        ListaDocumentosCartaXSLT contenid = new ListaDocumentosCartaXSLT() { Aplicacion = "CartasMilenia", RutaPlantillaXSLT = rutaXSLT, DatosXSLT = FiltrarXML(contenido.ToString()), NoCopias = 1 };
                        retorno.Add(contenid);
                        contadorMilenia++;



                    }

                }

                 return retorno;

            }
            catch (Exception ex)
            {
               // ExceptionPolicy.HandleException(ex, "Default"); //No se muestra la excepción para no detener las otras impresiones.

                return retorno;
            }
            return retorno;
        }
        private string FiltrarXML(string cadena)
        {
            cadena = cadena.Replace('"', ' ');
            cadena = cadena.Replace('&', ' ');
            return cadena;
        }
        private StringBuilder GetXmlPoliza(int estilo,
            string tipo,
            string poliza,
            string nombreCliente,
            string direccionCliente,
            string telefonoCliente,
            string garantiaExtendida,
            DateTime inicioVigencia,
            DateTime finVigencia,
            string articuloDescripcion,
            string esAplicaImpresionSerie,
            string numeroSerie,
            DateTime nuevasClausulas,
            bool printOld, string ip, out string rutaXSLT)
        {
            StringBuilder sb_contenido = null;

            rutaXSLT = "";
            string rutaWeb = "http://" + ip;
           // rutaXSLT = "http://" + ip + "/ElektraFront/Xsl/DetalleProductoCentral.xslt";
            switch (tipo)
            {
                case "MILTARREM"://reemplazo
                    /* Caducado 2014
                    if(printOld)
                    {
                        //sb_contenido = XmlMileniaSingleOld("/ElektraFront/Xsl/PolizaMileniaRemplazo.xslt",poliza,nombreCliente,direccionCliente,telefonoCliente,garantiaExtendida,inicioVigencia,finVigencia,articuloDescripcion, esAplicaImpresionSerie,numeroSerie,nuevasClausulas);
                    }
                    else
                    */
                    switch (estilo)
                    {
                        case 1:
                            //Poliza anterior al 2014 (Poliza Viejita)
                            /*
                                \\\\10.54.28.41\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\DetalleProductoCentral.xslt
                                http://10.54.28.41/ElektraFront/Xsl/DetalleProductoCentral.xslt
                              
                             */

                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaRemplazoEpos.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaRemplazo1.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                        case 2:
                            //Poliza después del 2014
                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaRemplazo2014.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaRemplazo2014.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                        case 3:
                            //Poliza que contiene una nueva clausula(Asterisco)
                            
                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaRemplazo2014Asterisco.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaRemplazo2014Asterisco.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                    }

                    break;



                case "MILCOMREM"://computo
                    /*if(printOld)
                    {
                        //sb_contenido = XmlMileniaSingleOld("/ElektraFront/Xsl/PolizaMileniaComputadoras.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie,nuevasClausulas);
                    }
                    else*/
                    switch (estilo)
                    {
                        case 1:
                            //Poliza anterior al 2014 (Poliza Viejita)
                            
                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaComputoEpos.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaComputo1.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                        case 2:
                            //Poliza después del 2014
                            
                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaComputo2014.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaComputo2014.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                        case 3:
                            //Poliza que contiene una nueva clausula(Asterisco)
                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaComputo2014Asterisco.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaComputo2014Asterisco.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                    }

                    break;




                case "MILMUEREM"://muebles
                    /*
                    if(printOld)
                    {
                        //sb_contenido = XmlMileniaSingle("/ElektraFront/Xsl/PolizaMileniaMuebles.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie,nuevasClausulas);
                    }
                    else
                    {*/
                    switch (estilo)
                    {
                        case 1:
                            //Poliza anterior al 2014 (Poliza Viejita)
                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaMueblesEpos.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaMuebles1.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                        case 2:
                            //Poliza después del 2014
                            rutaXSLT = "\\\\"+ip+"\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaMuebles2014.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaMuebles2014.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                        case 3:
                            //Poliza que contiene una nueva clausula(Asterisco)
                            rutaXSLT = "\\\\" + ip + "\\puntodeventa\\Tienda\\ElektraFront\\Xsl\\PolizaMileniaMuebles2014Asterisco.xslt";
                            sb_contenido = XmlMileniaSingle(rutaWeb+"/ElektraFront/Xsl/PolizaMileniaMuebles2014Asterisco.xslt", poliza, nombreCliente, direccionCliente, telefonoCliente, garantiaExtendida, inicioVigencia, finVigencia, articuloDescripcion, esAplicaImpresionSerie, numeroSerie, nuevasClausulas);
                            break;
                    }
                    break;
            }

            return sb_contenido;
        }
        public ArrayList ImprimeGarantiaMilenia(DataSet detalleMilenia,string ip)
        {
            ArrayList Retorno = new ArrayList();
            EntMileniaPosVenta ent = new EntMileniaPosVenta();
            try
            {
                
                if (detalleMilenia != null && detalleMilenia.Tables.Count > 0 && detalleMilenia.Tables[0].Rows.Count > 0)
                {
                    
                    //bool esNuevaMilenia = Convert.ToBoolean(((Hashtable)ConfigurationManager.Read("MileniaConfiguracion"))["NuevosCertificados"]);
                    //if (!esNuevaMilenia)
                    //{
                    //    Retorno = this.ImprimeGarantiaExtendida(impresionSurtimiento);
                    //}
                    //else
                    //{
                        if (ValidaImpresionMilenia())
                        {
                            if (this.ValidarExisteMilenia((int)detalleMilenia.Tables[0].Rows[0]["fiNoPedido"]))
                                Retorno = ImprimeGarantiaLigadaXML(detalleMilenia,ip);
                        }
                    //}
                }
                else
                {
                    System.Diagnostics.Trace.WriteLine("El pedido no contiene Milenias", "LOG");
                    Retorno = null;
                }
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("Error al Generar el XML de La garantia Milenia Ligada, En APIMILENIA", "LOG");
                throw new ApplicationException("No se pudo imprimir la Garantia Milenia por el siguiente motivo: " + ex.Message, ex.InnerException);
            }
            return Retorno;

        }


        public List<ListaDocumentosCartaXSLT> cadenaImpresionMilenia(int pedido, string ip)
        {
            List<ListaDocumentosCartaXSLT> cartas = new List<ListaDocumentosCartaXSLT>();
            EntMileniaPosVenta ent = new EntMileniaPosVenta();

            try {

                var det = ent.obtenerDetalleMilenias(pedido);


                ArrayList milenias = this.ImprimeGarantiaMilenia(det,ip);
                
                if (milenias == null) {
                    return cartas;
                }
                    

                for (int i = 0; i < milenias.Count; i++)
                {
                    cartas.Add((ListaDocumentosCartaXSLT)milenias[i]);
                    //sourcekey = "PolizasMilenia" + i.ToString();
                    
                }

                
            }
            catch (Exception e) { 
                   
            }

            return cartas;
        }
             


        #endregion





    }
}
