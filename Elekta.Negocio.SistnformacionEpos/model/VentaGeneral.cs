using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace Elekta.Negocio.SistnformacionEpos
{
    [DataContract]
    public class VentaGeneral
    {
        [DataMember]
        public int marcadas { get; set; }
        [DataMember]
        public int surtidas { get; set; }
        [DataMember]
        public int canceladasADE { get; set; }
        [DataMember]
        public int canceladasDDE { get; set; }
    }
}
