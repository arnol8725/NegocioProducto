using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;


namespace Elekta.Negocio.SistnformacionEpos
{
    [DataContract]
    public class ExistenciaInventario
    {
        [DataMember]
        public int ubicacion { get; set; }
        [DataMember]
        public string descripcion { get; set; }
        [DataMember]
        public int existencia { get; set; }
    }

    [DataContract]
    public class ExistenciaInventarioxTipo
    {
        [DataMember]
        public int Contadosinsurtir { get; set; }
        [DataMember]
        public int Apartadocon90pagado { get; set; }
        [DataMember]
        public int Apartadoconmenosde90pagado { get; set; }
        [DataMember]
        public int Creditocon90deenganchepagado { get; set; }
    }
}
