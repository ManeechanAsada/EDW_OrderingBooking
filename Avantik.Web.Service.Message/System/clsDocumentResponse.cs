using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Message.System
{
    [MessageContract]
    public class DocumentResponse : ResponseBase
    {
        [MessageBodyMember]
        public IEnumerable<DocumentView> DocumentType { get; set; }
    }
}
