using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Model;
using Avantik.Web.Service.Model.Contract;
using Avantik.Web.Service.Exception.Flight;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Model.Booking;
using Avantik.Web.Service.Extension;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity.Agency;
using Avantik.Web.Service.Message.Fee;
using Avantik.Web.Service.Message.SeatMap;
using Avantik.Web.Service.Exception.Booking;
using System.Web.Configuration;
using System.ServiceModel;

namespace Avantik.Web.Service
{
  [ServiceBehavior(AddressFilterMode = AddressFilterMode.Any)]

    public class AuthenticationService : IAuthenticationService
    {
        public InitializeResponse ServiceInitialize(InitializeRequest Request)
        {
            IBookingModelService objBookingService = BookingServiceFactory.CreateInstance();
            IAgencyService objAgencyService = AgencyServiceFactory.CreateInstance();
            InitializeResponse response = new InitializeResponse();
            string strHashing = string.Empty;
            string strEncryptKey = ConfigHelper.ToString("strKey");
            string serverName = ConfigHelper.ToString("ServerName");

            // set server id
            response.ServerId = serverName;

            // log in at dbo.user_account table
            Avantik.Web.Service.Message.TravelAgentLogonResponse travelAgentLogonResponse = GetUserLogon(Request.AgencyCode, Request.AgentLogon, Request.AgentPassword);

            if (travelAgentLogonResponse.Success == true)
            {
                string userId = travelAgentLogonResponse.AgentResponse.UserAccountId.ToString();
                string travelAgentCode = travelAgentLogonResponse.AgentResponse.AgencyCode.ToString();
                string cureencyRcd = travelAgentLogonResponse.AgentResponse.CurrencyRcd;
                string languageRcd = travelAgentLogonResponse.AgentResponse.LanguageRcd;

                Message.Agency.Agent agent = new Message.Agency.Agent();
                Byte checkAgency = 0;

                agent = travelAgentLogonResponse.AgentResponse;//objAgencyService.GetAgencySessionProfile(Request.AgencyCode, userId);

                if (agent.WebAgencyFlag == 1)
                {
                    if (agent.APIFlag == 1)
                    {
                        checkAgency = agent.APIFlag;

                        DateTime dtCurrentTime = DateTime.Now;
                        string strTime = dtCurrentTime.ToString("yyyy-MM-dd HH:mm");

                        // param combined with userid agencycode seat service passInfo flight cancel and current time
                        string strParams = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}", new string[] 
                                           { userId, travelAgentCode,agent.B2BAllowSeat.ToString(),agent.B2BAllowServices.ToString(), 
                                            agent.B2BAllowPassengerInfo.ToString(),agent.B2BAllowChangeFlight.ToString(), 
                                            agent.B2BAllowCancelFlight.ToString(),agent.B2BAllowNameChange.ToString(),agent.B2BAllowChangeDetail.ToString(),cureencyRcd,languageRcd,strTime });

                        strHashing = SecurityHelper.EncryptString(strParams, strEncryptKey);

                        if (!string.IsNullOrEmpty(strHashing))
                        {
                            response.Code = "000";
                            response.Message = "Success";
                            response.Success = true;
                            response.Token = strHashing;
                        }
                        else
                        {
                            response.Code = "A006";
                            response.Message = "Authenticate web service failed.";
                            response.Success = false;
                            response.Token = string.Empty;
                        }
                    }
                    else
                    {
                        response.Code = "A007";
                        response.Message = "Agency is not allowed use API.";
                        response.Success = false;
                        response.Token = string.Empty;
                    }
                }
                else
                {
                    response.Code = "A010";
                    response.Message = "Agency is not allowed use web.";
                    response.Success = false;
                    response.Token = string.Empty;
                }
            }
            else
            {
                response.Code = "A006";
                response.Message = "Authenticate web service failed.";
                response.Success = false;
                response.Token = string.Empty;
            }

            return response;
        }

        private Avantik.Web.Service.Message.TravelAgentLogonResponse GetUserLogon(string agencyCode, string userLogon, string password)
        {
            Avantik.Web.Service.Message.TravelAgentLogonRequest reqTravelAgentLogon = new Web.Service.Message.TravelAgentLogonRequest();
            Avantik.Web.Service.Message.TravelAgentLogonResponse travelAgentLogonResponse = new Web.Service.Message.TravelAgentLogonResponse();

            try
            {
                BookingService objBooking = new BookingService();

                reqTravelAgentLogon.AgencyCode = agencyCode;
                reqTravelAgentLogon.AgentLogon = userLogon;
                reqTravelAgentLogon.AgentPassword = password;

                travelAgentLogonResponse = objBooking.TravelAgentLogon(reqTravelAgentLogon);
            }
            catch
            {
                travelAgentLogonResponse.Success = false;
                travelAgentLogonResponse.Message = "User logon error.";
            }
            return travelAgentLogonResponse;

        }

    }
}
