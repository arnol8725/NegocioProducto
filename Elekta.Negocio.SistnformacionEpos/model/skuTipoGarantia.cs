using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;


namespace Elekta.Negocio.SistnformacionEpos
{
    [DataContract]
    public class skuTipoGarantia
    {
        [DataMember]
        public double sku { get; set; }
        [DataMember]
        public string tipoGarantia{ get; set; }
        [DataMember]
        public int cantidada{ get; set; }
    }
}
