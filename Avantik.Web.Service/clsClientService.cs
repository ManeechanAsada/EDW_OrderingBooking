using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Entity.Route;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Model.Contract;
using Avantik.Web.Service.Model.Factory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Avantik.Web.Service.Extension;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Model;
using Avantik.Web.Service.Message.Client;
using Avantik.Web.Service.Exception.Booking;
using System.ServiceModel;
namespace Avantik.Web.Service
{
                [ServiceBehavior(AddressFilterMode = AddressFilterMode.Any)]

    public class ClientService : Contracts.IClientService
    {
        public CreateClientProfileResponse CreateClientProfile(CreateClientProfileRequest request)
        {
            Model.Contract.IClientService objClientService = ClientServiceFactory.CreateInstance();
            CreateClientProfileResponse response = new CreateClientProfileResponse();
            Entity.Client.ClientProfile clientProfile = new Entity.Client.ClientProfile();
            Guid clientProfileId = Guid.NewGuid();
            bool result = false;

            // map request to entity
            clientProfile.Client = request.ClientProfile.Client.ToEntityClient(clientProfileId);
            clientProfile.BookingRemarks = null;
            clientProfile.PassengerProfiles = request.ClientProfile.PassengerProfiles.ToListEntityClient(clientProfileId);

            try
            {
                result = objClientService.ClientSave(clientProfile);

                if (result)
                {
                    // read
                    Entity.Client.ClientProfile clientProfileResult = objClientService.ClientRead(clientProfile.Client.ClientProfileId.ToString());
                    
                    if (clientProfileResult.Client != null)
                    {
                        response.ClientProfileId = clientProfileId;
                        response.ClientNumber = clientProfileResult.Client.ClientNumber;
                        response.Success = true;
                        response.Message = "Success";
                    }
                }
                else
                {
                    response.Success = false;
                    response.Message = "FAIL";
                }

            }
            catch
            {
                response.Success = false;
                response.Message = "FAIL";
            }
           
            return response;
        }

        public EditClientProfileResponse EditClientProfile(EditClientProfileRequest request)
        {
            Model.Contract.IClientService objClientService = ClientServiceFactory.CreateInstance();
            EditClientProfileResponse response = new EditClientProfileResponse();
            Entity.Client.ClientProfile clientProfile = new Entity.Client.ClientProfile();
            Guid clientProfileId = request.ClientProfile.Client.ClientProfileId;
            bool result = false;

            // map request to entity
            clientProfile.Client = request.ClientProfile.Client.ToEntityClient(clientProfileId);
            clientProfile.BookingRemarks = null;
            clientProfile.PassengerProfiles = request.ClientProfile.PassengerProfiles.ToListEntityClient(clientProfileId);

            try
            {
                result = objClientService.EditClientProfile(clientProfile);

                if (result)
                {
                    response.Success = true;
                    response.Message = "SUCCESS";
                }
                else
                {
                    response.Success = false;
                    response.Message = "FAIL";
                }

            }
            catch
            {
                response.Success = false;
                response.Message = "FAIL";
            }

            return response;
        }

        public CreateClientProfileResponse AddPassengerProfile(CreateClientProfileRequest request)
        {
            Model.Contract.IClientService objClientService = ClientServiceFactory.CreateInstance();
            CreateClientProfileResponse response = new CreateClientProfileResponse();
            Entity.Client.ClientProfile clientProfile = new Entity.Client.ClientProfile();
            Guid clientProfileId = request.ClientProfile.Client.ClientProfileId;
            bool result = false;

            // map request to entity
            clientProfile.Client = request.ClientProfile.Client.ToEntityClient(clientProfileId);
            clientProfile.BookingRemarks = null;
            clientProfile.PassengerProfiles = request.ClientProfile.PassengerProfiles.ToListEntityClient(clientProfileId);
            
            try
            {
                result = objClientService.ClientSave(clientProfile);

                if (result)
                {
                    response.ClientProfileId = clientProfileId;
                    response.Success = true;
                    response.Message = "SUCCESS";
                }
                else
                {
                    response.Success = false;
                    response.Message = "FAIL";
                }

            }
            catch
            {
                response.Success = false;
                response.Message = "FAIL";
            }

            return response;
        }

        public EditClientProfileResponse EditPassengerProfile(EditClientProfileRequest request)
        {
            Model.Contract.IClientService objClientService = ClientServiceFactory.CreateInstance();
            EditClientProfileResponse response = new EditClientProfileResponse();
            Entity.Client.ClientProfile clientProfile = new Entity.Client.ClientProfile();
            Guid clientProfileId = request.ClientProfile.Client.ClientProfileId;
            bool result = false;

            // map request to entity
            clientProfile.Client = request.ClientProfile.Client.ToEntityClient(clientProfileId);
            clientProfile.BookingRemarks = null;
            clientProfile.PassengerProfiles = request.ClientProfile.PassengerProfiles.ToListEntityClient(clientProfileId);
        
            try
            {
                result = objClientService.EditClientProfile(clientProfile);

                if (result)
                {
                    response.Success = true;
                    response.Message = "SUCCESS";
                }
                else
                {
                    response.Success = false;
                    response.Message = "FAIL";
                }

            }
            catch
            {
                response.Success = false;
                response.Message = "FAIL";
            }

            return response;
        }

        public AvailabilityResponse GetAvailability(AvailabilityRequest request)
        {
            AvailabilityResponse response = new AvailabilityResponse();
            AvailabilityService obj = new AvailabilityService();
            string userId = string.Empty;
            string agencyCode = string.Empty;
            string currencyCode = string.Empty;
            string languageCode = string.Empty;

            try
            {
                if (!string.IsNullOrEmpty(request.Token))
                {
                    // valid token
                    Avantik.Web.Service.Entity.Authentication objAuthen = Avantik.Web.Service.Infrastructrue.Authentication.Authenticate(request.Token);

                    if (objAuthen.ResponseSuccess)
                    {
                        //set user id  and agency code from obj authen
                        userId = objAuthen.UserId.ToString();
                        agencyCode = objAuthen.AgencyCode;
                        request.Token = request.Token;

                        // validate request
                        response = obj.GetAvailability(request);
                    }
                    else
                    {
                        response.Code = "A006";
                        response.Message = "Athenticate failed.";
                        response.Success = false;
                    }
                }
                else
                {
                    response.Code = "A004";
                    response.Message = "Require Token";
                    response.Success = false;
                }
            }
            catch (ModifyBookingException mex)
            {
                response.Code = mex.ErrorCode;
                response.Message = mex.Message;
                response.Success = false;
                //Logger.SaveLog("GetAvailability", DateTime.Now, DateTime.Now, mex.Message, XMLHelper.Serialize(request, false));
            }
            catch (System.Exception ex)
            {
                response.Code = "E001";
                response.Message = ex.Message;
                response.Success = false;
                // Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(GetSpecialServicesRequest), request));
                //Logger.SaveLog("GetAvailability", DateTime.Now, DateTime.Now, ex.Message, XMLHelper.Serialize(request, false));
            }

            return response;
        }

        // client logon

        // add flight

        // payment  cc  vc
    }
}
