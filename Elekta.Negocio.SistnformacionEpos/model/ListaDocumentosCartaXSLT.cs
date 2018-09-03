using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
namespace Elekta.Negocio.SistnformacionEpos
{
    [DataContract]
    public class ListaDocumentosCartaXSLT
    {
        [DataMember]
        public string Aplicacion { get; set; }
        [DataMember]
        public string RutaPlantillaXSLT { get; set; }
        [DataMember]
        public string DatosXSLT { get; set; }
        [DataMember]
        public int NoCopias { get; set; }


        
    }
}
