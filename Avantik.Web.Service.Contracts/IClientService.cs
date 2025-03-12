using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Contracts
{
    [ServiceContract(Namespace = "Avantik.Web.Service/")]
    public interface IClientService
    {
        //[OperationContract()]
        //CreateClientProfileResponse CreateClientProfile(CreateClientProfileRequest Request);

        //[OperationContract()]
        //EditClientProfileResponse EditClientProfile(EditClientProfileRequest request);

        //[OperationContract()]
        //CreateClientProfileResponse AddPassengerProfile(CreateClientProfileRequest request);

        //[OperationContract()]
        //EditClientProfileResponse EditPassengerProfile(EditClientProfileRequest request);

    }
}
