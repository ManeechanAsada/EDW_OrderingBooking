using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Model;
using Avantik.Web.Service.Model.Contract;
using Avantik.Web.Service.Model.Factory;
using Avantik.Web.Service.Exception.Flight;
using Avantik.Web.Service.Extension;
using System.ServiceModel;

namespace Avantik.Web.Service
{
     [ServiceBehavior(AddressFilterMode = AddressFilterMode.Any)]
    public class AvailabilityService : Service.Contracts.IAvailabilityService
    {
        public AvailabilityResponse GetAvailability(AvailabilityRequest Request)
        {
            AvailabilityResponse respose = new AvailabilityResponse();

            // valid token
            Avantik.Web.Service.Entity.Authentication objAuthen = Infrastructrue.Authentication.Authenticate(Request.Token);
            if (objAuthen.ResponseSuccess == false)
            {
                respose.Success = objAuthen.ResponseSuccess;
                respose.Message = objAuthen.ResponseMessage;
                respose.Code = objAuthen.ResponseCode;
                return respose;
            }
            else
            {
                try
                {
                    if (string.IsNullOrEmpty(Request.OriginRcd))
                    {
                        respose.Code = "V002";
                        respose.Success = false;
                        respose.Message = "OriginRcd parameter is required.";
                    }
                    else if (string.IsNullOrEmpty(Request.DestinationRcd))
                    {
                        respose.Code = "V003";
                        respose.Success = false;
                        respose.Message = "DestinationRcd parameter is required.";
                    }
                    else if (Request.FromDate == DateTime.MinValue)
                    {
                        respose.Code = "V004";
                        respose.Success = false;
                        respose.Message = "FromDate parameter is required.";
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(Request.AgencyCode))
                        {
                            Request.AgencyCode = objAuthen.AgencyCode;
                        }

                        Model.Contract.IAvailabilityService objInventoryService = FlightServiceFactory.CreateAvailabilityInstance(Request.AvailabilityTypes);

                        Guid flightId = Guid.Empty;
                        Guid fareId = Guid.Empty;

                        if (string.IsNullOrEmpty(Request.FlightId) == false)
                        {
                            flightId = new Guid(Request.FlightId);
                        }

                        if (string.IsNullOrEmpty(Request.FareId) == false)
                        {
                            fareId = new Guid(Request.FareId);
                        }

                        //Set default value.
                        if (string.IsNullOrEmpty(Request.CurrencyCode))
                        {
                            Request.CurrencyCode = "EN";
                        }
                        if (Request.ToDate == DateTime.MinValue)
                        {
                            Request.ToDate = Request.FromDate;
                        }

                        //Call get flight availability.
                        Entity.Flight.Availabilities availabilityResult = objInventoryService.GetAvailability(Request.OtherPassengerType,
                                                                                                              Request.BoardingClass,
                                                                                                              Request.BookingClass,
                                                                                                              Request.DayTimeIndicator,
                                                                                                              Request.ReturnDayTimeIndicator,
                                                                                                              Request.AgencyCode,
                                                                                                              Request.CurrencyCode,
                                                                                                              Request.Transitpoint,
                                                                                                              Request.PromotionCode,
                                                                                                              Request.FareTypes,
                                                                                                              Request.LanguageCode,
                                                                                                              Request.IpAddress,
                                                                                                              Request.OriginRcd,
                                                                                                              Request.DestinationRcd,
                                                                                                              Request.OdOriginRcd,
                                                                                                              Request.OdDestinationRcd,
                                                                                                              Request.FromDate,
                                                                                                              Request.ToDate,
                                                                                                              Request.ReturnFromDate,
                                                                                                              Request.ReturnToDate,
                                                                                                              Request.BookingDate,
                                                                                                              Request.MaxAmount,
                                                                                                              Request.Adult,
                                                                                                              Request.Child,
                                                                                                              Request.Infant,
                                                                                                              Request.Other,
                                                                                                              flightId,
                                                                                                              fareId,
                                                                                                              Request.NoneStopOnly,
                                                                                                              Request.IncludeDeparted,
                                                                                                              Request.IncludeCancelled,
                                                                                                              Request.IncludeWaitlisted,
                                                                                                              Request.IncludeSoldOut,
                                                                                                              Request.IncludeFares,
                                                                                                              Request.Refundable,
                                                                                                              Request.ReturnRefundable,
                                                                                                              Request.GroupFares,
                                                                                                              Request.ITFaresOnly,
                                                                                                              Request.StaffFares,
                                                                                                              Request.ApplyFareLogic,
                                                                                                              Request.ShowClose,
                                                                                                              0,
                                                                                                              Request.MapWithFares,
                                                                                                              Request.NoVat,
                                                                                                              Request.TransactionReference);


                        if (availabilityResult != null)
                        {
                            respose.Code = "000";
                            respose.Success = true;
                            respose.Message = "Success request.";

                            if (availabilityResult.FlightAvailabilityOutbound != null)
                            {
                                respose.AvailabilityOutbound = availabilityResult.FlightAvailabilityOutbound.MappingAvailabilityView();
                            }
                            if (availabilityResult.FlightAvailabilityReturn != null)
                            {
                                respose.AvailabilityReturn = availabilityResult.FlightAvailabilityReturn.MappingAvailabilityView();
                            }
                        }
                        else
                        {
                            respose.Code = "V001";
                            respose.Success = false;
                            respose.Message = "No Availability result return.";
                        }

                    }
                }
                catch (AvailabilityException ex)
                {
                    respose.Code = "E001";
                    respose.Success = false;
                    respose.Message = ex.Message;
                }
                catch (System.Exception ex)
                {
                    respose.Code = "E001";
                    respose.Message = ex.Message;
                    respose.Success = false;
                    //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(AvailabilityRequest), Request));
                }
            }
            return respose;
        }
    }
}
