using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Message.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Contracts
{
    [ServiceContract(Namespace = "Avantik.Web.Service/")]
    public interface IAuthenticationService
    {
       // [OperationContract()]
       // InitializeResponse ServiceInitialize(InitializeRequest Request);

    }
}
