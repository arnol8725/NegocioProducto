using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;


namespace Elekta.Negocio.SistnformacionEpos
{
    [DataContract]
    public class DetalleProducto:Respuesta
    {
        [DataMember]
        public List<DetalleBusqueda> detBusquesa { get; set; }
    }
}
