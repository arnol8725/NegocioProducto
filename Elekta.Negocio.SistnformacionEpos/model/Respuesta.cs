using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
namespace Elekta.Negocio.SistnformacionEpos
{
    [DataContract]
    public class Respuesta
    {
        [DataMember]
        public bool error { get; set; }

        [DataMember]
        public string mensaje{ get; set; }

        [DataMember]
        public string mensajeTecnico { get; set; }
    }
}
