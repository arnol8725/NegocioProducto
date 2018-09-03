using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace Elekta.Negocio.SistnformacionEpos
{

    [DataContract]
    public class DetalleBusqueda
    {
        [DataMember]
        public double codigo{get;set;}
        [DataMember]
        public string descripcion{get;set;}
        [DataMember]
        public string marca { get; set; }
        [DataMember]
        public string precio { get; set; }
    }
}
