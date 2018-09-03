using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace Elekta.Negocio.SistnformacionEpos
{
    [DataContract]
    public class DetalleProductoxSKU:Respuesta
    {
        [DataMember]
        public string codigo { get; set; }
        [DataMember]
        public string descripcion { get; set; }
        [DataMember]
        public string negocio { get; set; }
        [DataMember]
        public string depto { get; set; }
        [DataMember]
        public string subDepto { get; set; }
        [DataMember]
        public string clase { get; set; }
        [DataMember]
        public string subClase { get; set; }
        [DataMember]
        public string tipoProducto { get; set; }
        [DataMember]
        public string estado { get; set; }
        [DataMember]
        public string fechaAlta  { get; set; }
        [DataMember]
        public string precioContado { get; set; }
        [DataMember]
        public string descMaximo { get; set; }
        [DataMember]
        public string ultCambioPrecio { get; set; }
        [DataMember]
        public string tiempoGarantia { get; set; }        
        [DataMember]
        public string diasReparacion { get; set; }
        [DataMember]
        public string requerido { get; set; }

        [DataMember]
        public string paquete { get; set; }

        [DataMember]
        public string manejaNoSerie { get; set; }
        [DataMember]
        public string codigoBarras { get; set; }
        [DataMember]
        public string tipoBloqueo { get; set; }
        [DataMember]
        public List<DetalleMovimiento> detMov { get; set; }
        [DataMember]
        public List<ExistenciaInventario> detExistInv { get; set; }
        [DataMember]
        public List<VentaGeneral> detVtaGeneral { get; set; }

        [DataMember]
        public List<ExistenciaInventarioxTipo> detInvTipo { get; set; }
        

    }
}
