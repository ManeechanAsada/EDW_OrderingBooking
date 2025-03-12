using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace Avantik.Web.Service.Proxy
{
    public class ClientServiceProxy : ClientBase<IClientService>, IClientService
    {
        //public CreateClientProfileResponse CreateClientProfile(CreateClientProfileRequest request)
        //{
        //    return base.Channel.CreateClientProfile(request);
        //}
        //public EditClientProfileResponse EditClientProfile(EditClientProfileRequest request)
        //{
        //    return base.Channel.EditClientProfile(request);
        //}
        //public CreateClientProfileResponse AddPassengerProfile(CreateClientProfileRequest request)
        //{
        //    return base.Channel.AddPassengerProfile(request);
        //}
        //public EditClientProfileResponse EditPassengerProfile(EditClientProfileRequest request)
        //{
        //    return base.Channel.EditPassengerProfile(request);
        //}

    }
}
