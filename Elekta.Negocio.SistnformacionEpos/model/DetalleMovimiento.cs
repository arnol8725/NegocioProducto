using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;


namespace Elekta.Negocio.SistnformacionEpos
{
    
    [DataContract]
    public class DetalleMovimiento
    {
        [DataMember]
        public double NoTransac { get; set; }
        [DataMember]
        public string UbiId { get; set; }
        [DataMember]
        public string DMInvCant { get; set; }
        [DataMember]
        public string TipoOp { get; set; }  
        [DataMember]
        public string TopDesc { get; set; }  
        [DataMember]
        public string MInvRef { get; set; }  
        [DataMember]
        public string ExistIni { get; set; }  
        [DataMember]
        public string ExistFin { get; set; }  
        [DataMember]
        public string MInvEntSal { get; set; }  
        [DataMember]
        public string EmpIDRed { get; set; }  
        [DataMember]
        public string WS { get; set; } 
        [DataMember]
        public string UbiDesc { get; set; } 
        [DataMember]
        public string fecha { get; set; } 

    }
}
