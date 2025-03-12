using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Entity.Client;

namespace Avantik.Web.Service.Model.Contract
{
    public interface IClientService
    {
        bool ClientSave(ClientProfile clientProfile);
        ClientProfile ClientRead(string clientProfileId);
        bool EditClientProfile(ClientProfile clientProfile);


    }
}
