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
using System.Web;
using System.ServiceModel;


namespace Avantik.Web.Service
{
   [ServiceBehavior(AddressFilterMode = AddressFilterMode.Any)]
    public class BookingService : IBookingService
    {
        public BookingSaveResponse SaveBooking(BookingSaveRequest Request)
        {
            IBookingModelService objBookingService = BookingServiceFactory.CreateInstance();
            BookingSaveResponse response = new BookingSaveResponse();
            BookingPaymentResponse paymentResponse = null;
            Booking booking = new Booking();
            bool result = false;
            bool createTickets = false;
            bool readBooking = false;
            bool readOnly = false;
            bool bSetLock = false;
            bool bCheckSeatAssignment = Request.CheckSeatAssignment;
            bool bCheckSessionTimeOut = Request.CheckSessionTimeOut;
            Guid userId = Guid.Empty; 
            string agencyCode = string.Empty;

            try
            {
                // valid token
                Avantik.Web.Service.Entity.Authentication objAuthen = Infrastructrue.Authentication.Authenticate(Request.Token);
                if (objAuthen.ResponseSuccess == false)
                {
                    response.Success = objAuthen.ResponseSuccess;
                    response.Message = objAuthen.ResponseMessage;
                    response.Code = objAuthen.ResponseCode;
                    return response;
                }
                else
                {
                    // prepare data
                    userId = objAuthen.UserId;
                    agencyCode = objAuthen.AgencyCode;

                    //Header
                    if (Request.Booking.Header != null)
                    {
                        Request.Booking.Header.UpdateBy = userId;
                        Request.Booking.Header.CreateBy = userId;
                        Request.Booking.Header.AgencyCode = agencyCode;
                    }

                    //Segment
                    if (Request.Booking.FlightSegments != null && Request.Booking.FlightSegments.Count > 0)
                    {
                        foreach (Message.Booking.FlightSegment obj in Request.Booking.FlightSegments)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                        }
                    }
                    //passenger
                    if (Request.Booking.Passengers != null && Request.Booking.Passengers.Count > 0)
                    {
                        foreach (Message.Booking.Passenger obj in Request.Booking.Passengers)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                        }
                    }
                    //mapping
                    if (Request.Booking.Mappings != null && Request.Booking.Mappings.Count > 0)
                    {
                        foreach (Message.Booking.Mapping obj in Request.Booking.Mappings)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                            obj.AgencyCode = agencyCode;
                        }
                    }
                    //fee
                    if (Request.Booking.Fees != null && Request.Booking.Fees.Count > 0)
                    {
                        foreach (Message.Booking.Fee obj in Request.Booking.Fees)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                            obj.AgencyCode = agencyCode;
                        }
                    }
                    //tax
                    if (Request.Booking.Taxs != null && Request.Booking.Taxs.Count > 0)
                    {
                        foreach (Message.Booking.Tax obj in Request.Booking.Taxs)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                        }
                    }
                    //remark
                    if (Request.Booking.Remarks != null && Request.Booking.Remarks.Count > 0)
                    {
                        foreach (Message.Booking.Remark obj in Request.Booking.Remarks)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                        }
                    }
                    //service
                    if (Request.Booking.Services != null && Request.Booking.Services.Count > 0)
                    {
                        foreach (Message.Booking.PassengerService obj in Request.Booking.Services)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                        }
                    }
                    //payment
                    if (Request.Booking.Payments != null && Request.Booking.Payments.Count > 0)
                    {
                        foreach (Message.Booking.Payment obj in Request.Booking.Payments)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                            obj.AgencyCode = agencyCode;
                        }
                    }

                }

                // map from message to entity
                if (Request != null)
                {
                    booking.Header = Request.Booking.Header.ToEntity();
                    booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                    booking.Passengers = Request.Booking.Passengers.ToListEntity();
                    booking.Mappings = Request.Booking.Mappings.ToListEntity();
                    booking.Payments = Request.Booking.Payments.ToListEntity();
                    booking.Fees = Request.Booking.Fees.ToListEntity();
                    booking.Remarks = Request.Booking.Remarks.ToListEntity();
                    booking.Services = Request.Booking.Services.ToListEntity();
                    booking.Taxs = Request.Booking.Taxs.ToListEntity();

                    //set flag agency
                    if (booking.Header != null)
                    {
                        Agent agent = new Agent();
                        IAgencyService objAgencyService = AgencyServiceFactory.CreateInstance();

                        agent = objAgencyService.GetAgencySessionProfile(agencyCode, string.Empty);

                        if (agent != null)
                        {
                            //set booking source
                            booking.SetBookingSource(agent.OwnAgencyFlag, agent.WebAgencyFlag);
                        }
                    }
                }
                    
                result = objBookingService.SaveBooking(booking,
                                                       createTickets,
                                                       readBooking,
                                                       readOnly,
                                                       bSetLock,
                                                       bCheckSeatAssignment,
                                                       bCheckSessionTimeOut);

                if (result == true && booking.Payments != null && booking.Payments.Count > 0)
                {
                    BookingPaymentRequest payRequest = new BookingPaymentRequest();
                    payRequest.Mappings = Request.Booking.Mappings;
                    payRequest.Payments = Request.Booking.Payments;
                    payRequest.Fees = Request.Booking.Fees;

                    paymentResponse = BookingPayment(payRequest);

                    if (paymentResponse.Success == false) result = false;
                }

                if (result == true)
                {
                    response.Success = true;
                    response.Message = "Success request.";
                }
                else
                {
                    response.Success = false;
                    response.Message = "No Booking Save";
                }
            }
            catch (BookingSaveException ex)
            {
                response.Success = false;
                response.Message = ex.Message;
                response.Code = ex.ErrorCode;
            }
            catch (System.Exception ex)
            {
                response.Success = false;
                response.Message = "Booking Save error.";
                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(BookingSaveRequest), Request));
            }

            return response;
        }

        public BookingReadResponse ReadBooking(BookingReadRequest Request)
        {
            IBookingModelService objBookingService = BookingServiceFactory.CreateInstance();
            BookingReadResponse response = new BookingReadResponse();

            string bookingReference = string.Empty;
            double bookingNumber = 0;
            bool bReadonly = false;
            bool bSeatLock = false;
            string bookingId = string.Empty;
            string bookingRef = string.Empty;
            string userId = string.Empty;
            bool bReadHeader = true;
            bool bReadSegment = true;
            bool bReadPassenger = true;
            bool bReadRemark = true;
            bool bReadPayment = true;
            bool bReadMapping = true;
            bool bReadService = true;
            bool bReadTax = true;
            bool bReadFee = true;
            bool bReadOd = true;
            string releaseBookingId = string.Empty;
            string CompleteRemarkId = string.Empty;

            try
            {
                // valid token
                Avantik.Web.Service.Entity.Authentication objAuthen = Avantik.Web.Service.Infrastructrue.Authentication.Authenticate(Request.Token);
                if (objAuthen.ResponseSuccess == false)
                {
                    response.Success = objAuthen.ResponseSuccess;
                    response.Message = objAuthen.ResponseMessage;
                    response.Code = objAuthen.ResponseCode;
                    return response;
                }

                if (String.IsNullOrEmpty(Request.BookingId) && String.IsNullOrEmpty(Request.BookingReference))
                {
                    response.Success = false;
                    response.Message = "Require Booking ID or Booking Reference.";
                    response.Code = "299";
                }
                else
                {
                    if (String.IsNullOrEmpty(Request.BookingId) && !String.IsNullOrEmpty(Request.BookingReference))
                    {
                        bookingId = Guid.Empty.ToString();
                        bookingRef = Request.BookingReference;
                    }
                    else if (!String.IsNullOrEmpty(Request.BookingId) && String.IsNullOrEmpty(Request.BookingReference))
                    {
                        bookingId = Request.BookingId;
                        bookingRef = string.Empty;
                    }
                    else
                    {
                        bookingId = Request.BookingId;
                        bookingRef = Request.BookingReference;
                    }
                    Booking booking = objBookingService.ReadBooking(bookingId,
                                                                         bookingRef,
                                                                         bookingNumber,
                                                                         bReadonly,
                                                                         bSeatLock,
                                                                         userId,
                                                                         bReadHeader,
                                                                         bReadSegment,
                                                                         bReadPassenger,
                                                                         bReadRemark,
                                                                         bReadPayment,
                                                                         bReadMapping,
                                                                         bReadService,
                                                                         bReadTax,
                                                                         bReadFee,
                                                                         bReadOd,
                                                                         releaseBookingId,
                                                                         CompleteRemarkId);

                    if (booking != null && booking.Header != null && booking.Segments != null && booking.Segments.Count > 0 &&
                        booking.Passengers != null && booking.Passengers.Count > 0 && booking.Mappings != null && booking.Mappings.Count > 0)
                    {
                        response.BookingResponse = new BookingResponse();
                        response.BookingResponse.Header = booking.Header.ToBookingMessage();
                        response.BookingResponse.FlightSegments = booking.Segments.ToBookingMessage();
                        response.BookingResponse.Passengers = booking.Passengers.ToBookingMessage();
                        response.BookingResponse.Mappings = booking.Mappings.ToBookingMessage();
                        response.BookingResponse.Payments = booking.Payments.ToBookingMessage();
                        response.BookingResponse.Remarks = booking.Remarks.ToBookingMessage();
                        response.BookingResponse.Fees = booking.Fees.ToBookingMessage();
                        response.BookingResponse.Taxs = booking.Taxs.ToBookingMessage();
                        response.BookingResponse.Services = booking.Services.ToBookingMessage();
                        response.BookingResponse.Quotes = booking.Quotes.ToBookingMessage();

                        response.BookingResponse.RecordLocator = null;
                        response.BookingResponse.BookingId = null;

                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }
                    else
                    {
                        response.Code = "B008";
                        response.Message = "Read Booking failed.";
                        response.Success = false;
                    }
                }
            }
            catch (System.Exception ex)
            {
                response.Code = "129";
                response.Message = "No PNR Match Found";
                response.Success = false;
                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(BookingReadRequest), Request));
            }

            return response;
        }

        public BookingPaymentResponse BookingPayment(BookingPaymentRequest Request)
        {
            IPaymentService objPaymentService = PaymentServiceFactory.CreateInstance();
            BookingPaymentResponse response = new BookingPaymentResponse();

            Booking booking = null;
            bool resultPayment = false;
            Guid userId = new Guid();
            IList<PaymentAllocation> paymentAllocations = null;

            try
            {
                // valid token
                Avantik.Web.Service.Entity.Authentication objAuthen = Avantik.Web.Service.Infrastructrue.Authentication.Authenticate(Request.Token);
                if (objAuthen.ResponseSuccess == false)
                {
                    response.Success = objAuthen.ResponseSuccess;
                    response.Message = objAuthen.ResponseMessage;
                    response.Code = objAuthen.ResponseCode;
                    return response;
                }
                else
                {
                    userId = objAuthen.UserId;

                    //payment
                    if (Request.Payments != null && Request.Payments.Count > 0)
                    {
                        foreach (Message.Booking.Payment obj in Request.Payments)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                        }
                    }

                    //payment
                    if (Request.Fees != null && Request.Fees.Count > 0)
                    {
                        foreach (Message.Booking.Fee obj in Request.Fees)
                        {
                            obj.CreateBy = userId;
                            obj.UpdateBy = userId;
                        }
                    }


                }

                booking = new Booking();
                booking.Payments = Request.Payments.ToListEntity();
                booking.Mappings = Request.Mappings.ToListEntity();
                booking.Fees = Request.Fees.ToListEntity();
                paymentAllocations = booking.CreateAllocation();

                if (paymentAllocations != null)
                    resultPayment = objPaymentService.SavePayment(booking.Payments, paymentAllocations);
               
                if(resultPayment == true)
                {
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch (System.Exception ex)
            {
                response.Message = "Save payment error";
                response.Success = false;
                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(BookingPaymentRequest), Request));
            }

            return response;
        }

        public BookingPaymentCreditCardResponse BookingPaymentCreditCard(BookingPaymentCreditCardRequest Request)
        {
            IPaymentService objPaymentService = PaymentServiceFactory.CreateInstance();
            BookingPaymentCreditCardResponse response = new BookingPaymentCreditCardResponse();

            Booking booking = null;
            bool resultPayment = false;
            IList<PaymentAllocation> paymentAllocations = null;

            try
            {
                if (Request != null)
                {
                    booking = new Booking();
                    booking.Header = Request.Booking.Header.ToEntity();
                    booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                    booking.Passengers = Request.Booking.Passengers.ToListEntity();
                    booking.Remarks = Request.Booking.Remarks.ToListEntity();
                    booking.Payments = Request.Booking.Payments.ToListEntity();
                    booking.Mappings = Request.Booking.Mappings.ToListEntity();
                    booking.Services = Request.Booking.Services.ToListEntity();
                    booking.Taxs = Request.Booking.Taxs.ToListEntity();

                    booking.Fees = Request.Booking.Fees.ToListEntity();
                    booking.Fees.Add(Request.PaymentFee.ToEntity());

                    paymentAllocations = booking.CreateAllocation();

                    //resultPayment = objPaymentService.PaymentCreditCard(booking, paymentAllocations,
                    //                                                    Request.SecurityToken, Request.AuthenticationToken,
                    //                                                    Request.CommerceIndicator, Request.BookingReference,
                    //                                                    Request.RequestSource, true, false, false);
                    if (resultPayment == true)
                    {
                        response.Message = "Success";
                        response.Success = true;
                    }
                    else
                    {
                        response.Message = "PaymentCreditCard failed";
                        response.Success = false;
                    }
                }
                else
                {
                    response.Message = "BookingPaymentCreditCardRequest is null.";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "BookingPaymentCreditCard Error.";
                response.Success = false;
            }

            return response;
        }

        public BookingReadVoucherResponse BookingReadVoucher(BookingReadVoucherRequest Request)
        {
            IVoucherService objVoucherService = VoucherServiceFactory.CreateInstance();
            BookingReadVoucherResponse response = new BookingReadVoucherResponse();
            List<Entity.Voucher> vouchers = new List<Entity.Voucher>();
            Entity.Voucher voucher = null;
            bool bResult = true;

            try
            {
                if (Request != null)
                {
                    foreach (Message.Voucher v in Request.Vouchers)
                    {
                       // voucher = objVoucherService.GetVoucher(v.ToVoucherEntity());
                        if (voucher != null)
                        {
                            vouchers.Add(voucher);
                        }
                        else
                        {
                            bResult = false;
                        }
                    }

                    if (bResult)
                    {
                        response.Vouchers = vouchers.ToVoucherMessage();
                        response.Message = "Success";
                        response.Success = true;
                    }
                    else
                    {
                        response.Message = "Read voucher failed.";
                        response.Success = false;
                    }
                }
                else
                {
                    response.Message = "BookingReadVoucherRequest is null.";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "BookingReadVoucher error.";
                response.Success = false;
            }

            return response;
        }

        public BookingPaymentVoucherResponse BookingPaymentVoucher(BookingPaymentVoucherRequest Request)
        {
            IVoucherService objVoucherService  = VoucherServiceFactory.CreateInstance();
            IPaymentService objPaymentService = PaymentServiceFactory.CreateInstance();
            BookingPaymentVoucherResponse response = new BookingPaymentVoucherResponse();
            Entity.Voucher voucher = new Entity.Voucher();
            IList<Entity.Booking.Payment> payments = null;
            IList<PaymentAllocation> paymentAllocations = null;
            Booking booking = null;
            bool bResult = true;

            try
            {
                if (Request != null)
                {
                    //check duplicate
                    payments = Request.Booking.Payments.ToListEntity();
                    bResult = voucher.ValidateVoucherDuplicate(payments);
                    if (bResult)
                    {
                        foreach (Entity.Booking.Payment p in payments)
                        {
                            voucher.VoucherNumber = p.DocumentNumber;

                          //  voucher = objVoucherService.GetVoucher(voucher);

                            if (voucher == null)
                            {
                                bResult = false;
                            }
                        }

                        if (bResult)
                        {
                            if (voucher.ValidateVoucherPayment(payments) == true)
                            {
                                booking = new Booking();
                                booking.Payments = Request.Booking.Payments.ToListEntity();
                                booking.Mappings = Request.Booking.Mappings.ToListEntity();
                                booking.Fees = Request.Booking.Fees.ToListEntity();
                                paymentAllocations = booking.CreateAllocation();
                                bResult = objPaymentService.SavePayment(Request.Booking.Payments.ToListEntity(), paymentAllocations);

                                if (bResult)
                                {
                                    response.Message = "Success";
                                    response.Success = true;
                                }
                                else
                                {
                                    response.Message = "SavePayment failed.";
                                    response.Success = false;
                                }
                            }
                            else
                            {
                                bResult = false;
                                response.Message = "Vouchers amount not enough.";
                                response.Success = false;
                            }
                        }
                        else
                        {
                            response.Message = "Read voucher failed.";
                            response.Success = false;
                        }
                    }
                    else
                    {
                        response.Message = "Duplicate vouchers.";
                        response.Success = false;
                    }
                }
                else
                {
                    response.Message = "BookingPaymentVoucherRequest is null.";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "BookingPaymentVoucher error.";
                response.Success = false;
            }

            return response;
        }

        public BookingPaymentCreditAgencyResponse BookingPaymentCreditAgency(BookingPaymentCreditAgencyRequest Request)
        {
            IAgencyService objAgencyService = AgencyServiceFactory.CreateInstance();
            IPaymentService objPaymentService = PaymentServiceFactory.CreateInstance();
            BookingPaymentCreditAgencyResponse response = new BookingPaymentCreditAgencyResponse();
            IList<Entity.Booking.Payment> payments = null;
            IList<PaymentAllocation> paymentAllocations = null;
            Booking booking = null;
            Agent agent = new Agent();
            bool bResult = true;

            try
            {
                if (Request != null)
                {
                    agent = objAgencyService.GetAgencySessionProfile(Request.Booking.Header.AgencyCode, string.Empty);
                    if (agent != null)
                    {
                        payments = Request.Booking.Payments.ToListEntity();
                        if (payments != null)
                        {
                            foreach (Entity.Booking.Payment p in payments)
                            {
                                if (!p.ValidateCreditAgency(agent))
                                {
                                    bResult = false;
                                    response.Message = "ValidateCreditAgency failed.";
                                    response.Success = false;
                                }
                            }

                            if (bResult)
                            {
                                booking = new Booking();
                                booking.Payments = Request.Booking.Payments.ToListEntity();
                                booking.Mappings = Request.Booking.Mappings.ToListEntity();
                                booking.Fees = Request.Booking.Fees.ToListEntity();
                                paymentAllocations = booking.CreateAllocation();
                                bResult = objPaymentService.SavePayment(Request.Booking.Payments.ToListEntity(), paymentAllocations);

                                if (bResult)
                                {
                                    response.Message = "Success";
                                    response.Success = true;
                                }
                                else
                                {
                                    response.Message = "SavePayment failed.";
                                    response.Success = false;
                                }
                            }
                        }
                    }
                    else
                    {
                        response.Message = "GetAgencySessionProfile failed.";
                        response.Success = false;
                    }
                }
                else
                {
                    response.Message = "BookingPaymentCreditAgencyRequest is null.";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "BookingPaymentCreditAgency error.";
                response.Success = false;
            }
            return response;
        }

        public AgencySessionProfileResponse GetAgencySessionProfile(AgencySessionProfileRequest Request)
        {
            IAgencyService objAgencyService = AgencyServiceFactory.CreateInstance();
            AgencySessionProfileResponse response = new AgencySessionProfileResponse();
            Agent agent = new Agent();

            try
            {
                agent = objAgencyService.GetAgencySessionProfile(Request.AgencyCode, Request.UserId);

                if (agent != null)
                {
                    response.AgentResponse = agent.ToAgentLogonMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Get Agency Profile error.";
                response.Success = false;
            }

            return response;
        }

        public TravelAgentLogonResponse TravelAgentLogon(TravelAgentLogonRequest Request)
        {
            IAgencyService objAgencyService = AgencyServiceFactory.CreateInstance();
            TravelAgentLogonResponse response = new TravelAgentLogonResponse();
            Agent agent = new Agent();

            try
            {
                agent = objAgencyService.TravelAgentLogon(Request.AgencyCode, Request.AgentLogon, Request.AgentPassword);

                if (agent != null && agent.Users != null)
                {
                    response.AgentResponse = agent.ToAgentLogonMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch(System.Exception ex)
            {
                response.Message = "Agent logon error.";
                response.Success = false;
                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, "Travel Agent Logon");
            }

            return response;
        }

        public BookingFlightResponse BookFlight(BookingFlightRequest Request)
        {
            IBookingModelService objBookingService = BookingServiceFactory.CreateInstance();
            BookingFlightResponse response = new BookingFlightResponse();
            IList<Entity.Booking.Flight> flights = null;

            try
            {
                // valid token
                Avantik.Web.Service.Entity.Authentication objAuthen = Avantik.Web.Service.Infrastructrue.Authentication.Authenticate(Request.Token);

                if (objAuthen.ResponseSuccess == false)
                {
                    response.Success = objAuthen.ResponseSuccess;
                    response.Message = objAuthen.ResponseMessage;
                    response.Code = objAuthen.ResponseCode;
                    return response;
                }
                else
                {
                    Request.UserId = "{"+ objAuthen.UserId.ToString() + "}";

                    if (string.IsNullOrEmpty(Request.AgencyCode))
                        Request.AgencyCode = objAuthen.AgencyCode;
                }

                //map from message to entity
                if (Request != null)
                {
                    flights = Request.Flight.ToListEntity();
                }
                else
                {
                    response.Message = "Map message to entity failed.";
                    response.Success = false;
                }

                Booking booking = objBookingService.BookFlight(Request.AgencyCode,
                                                              Request.Currency,
                                                              flights,
                                                              Request.BookingId,
                                                              Request.Adults,
                                                              Request.Children,
                                                              Request.Infants,
                                                              Request.Others,
                                                              Request.StrOthers,
                                                              Request.UserId,
                                                              Request.StrIpAddress,
                                                              Request.StrLanguageCode,
                                                              Request.BNoVat);
                if (booking != null && booking.Header != null && booking.Segments != null && booking.Segments.Count > 0 &&
                    booking.Passengers != null && booking.Passengers.Count > 0 && booking.Mappings != null && booking.Mappings.Count > 0)
                {
                    response.BookFlightResponse = new BookFlightResponse();
                    response.BookFlightResponse.Header = booking.Header.ToBookingMessage();
                    response.BookFlightResponse.FlightSegments = booking.Segments.ToBookingMessage();
                    response.BookFlightResponse.Passengers = booking.Passengers.ToBookingMessage();
                    response.BookFlightResponse.Mappings = booking.Mappings.ToBookingMessage();
                    response.BookFlightResponse.Payments = booking.Payments.ToBookingMessage();
                    response.BookFlightResponse.Remarks = booking.Remarks.ToBookingMessage();
                    response.BookFlightResponse.Fees = booking.Fees.ToBookingMessage();
                    response.BookFlightResponse.Taxs = booking.Taxs.ToBookingMessage();
                    response.BookFlightResponse.Services = booking.Services.ToBookingMessage();
                    response.BookFlightResponse.Quotes = booking.Quotes.ToBookingMessage();

                    if (booking.FindUSSegment() == true)
                    {
                        response.Message = "Selected flight full";
                        response.Success = false;
                    }
                    else
                    {
                        response.Message = "Success";
                        response.Success = true;
                    }

                }
                else
                {
                    response.Message = "Booking error";
                    response.Success = false;
                }
            }
            catch (BookingException ex)
            {
                response.Message = ex.Message;
                response.Success = false;
            }
            catch(System.Exception ex)
            {
                response.Message = ex.Message;
                response.Success = false;
                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(BookingFlightRequest), Request));
            }
            return response;
        }

        public GetFeeResponse GetFee(GetFeeRequest Request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            GetFeeResponse response = new GetFeeResponse();
            try
            {
                IList<Entity.Fee> Fees = objFeeService.GetFee(Request.StrFeeRcd, 
                    Request.StrCurrencyCode, 
                    Request.StrAgencyCode, 
                    Request.StrClass, 
                    Request.StrFareBasis, 
                    Request.StrOrigin, 
                    Request.StrDestination, 
                    Request.StrFlightNumber,
                    Request.DtDate, 
                    Request.StrLanguage, 
                    Request.bNoVat);

                if (Fees != null && Fees.Count > 0)
                {
                    // map to response
                    response.Fees = Fees.ToFeeMessage();
                    response.Message = "Success";
                    response.Success = true;
                    response.Code = "000";

                }
                else
                {
                    response.Code = "V010";
                    response.Message = "Get Fee failed.";
                    response.Success = false;
                }
            }
            catch
            {
                response.Code = "V010";
                response.Message = "Get Fee failed.";
                response.Success = false;
            }

            return response;
        }
        
        public CalculateFeesBookingResponse CalculateFeesBookingCreate(CalculateFeesBookingCreateRequest Request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            CalculateFeesBookingResponse response = new CalculateFeesBookingResponse();

            Booking booking = new Booking();

            // map from message to entity
            if (Request != null)
            {
                booking.Header = Request.Booking.Header.ToEntity();
                booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                booking.Passengers = Request.Booking.Passengers.ToListEntity();
                booking.Mappings = Request.Booking.Mappings.ToListEntity();
                booking.Payments = Request.Booking.Payments.ToListEntity();
                booking.Fees = Request.Booking.Fees.ToListEntity();
                booking.Remarks = Request.Booking.Remarks.ToListEntity();
                booking.Services = Request.Booking.Services.ToListEntity();
                booking.Taxs = Request.Booking.Taxs.ToListEntity();
            }

            try
            {
                IList<Avantik.Web.Service.Entity.Booking.Fee> BookingFees = objFeeService.CalculateFeesBookingCreate(Request.AgencyCode, Request.Currency, booking, Request.Language);

                if (BookingFees != null)
                {
                    // map to response
                    response.Fees = BookingFees.ToBookingMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "CalculateFeesBookingCreate error";
                response.Success = false;
            }

            return response;
        }
        
        public CalculateFeesBookingResponse CalculateFeesBookingChange(CalculateFeesBookingCreateRequest Request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            CalculateFeesBookingResponse response = new CalculateFeesBookingResponse();

            Booking booking = new Booking();

            // map from message to entity
            if (Request != null)
            {
                booking.Header = Request.Booking.Header.ToEntity();
                booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                booking.Passengers = Request.Booking.Passengers.ToListEntity();
                booking.Mappings = Request.Booking.Mappings.ToListEntity();
                booking.Payments = Request.Booking.Payments.ToListEntity();
                booking.Fees = Request.Booking.Fees.ToListEntity();
                booking.Remarks = Request.Booking.Remarks.ToListEntity();
                booking.Services = Request.Booking.Services.ToListEntity();
                booking.Taxs = Request.Booking.Taxs.ToListEntity();
            }

            try
            {
                IList<Avantik.Web.Service.Entity.Booking.Fee> BookingFees = objFeeService.CalculateFeesBookingChange(Request.AgencyCode, Request.Currency, booking, Request.Language);

                if (BookingFees != null)
                {
                    // map to response
                    response.Fees = BookingFees.ToBookingMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "CalculateFeesBookingChange error";
                response.Success = false;
            }

            return response;
        }
        
        public CalculateFeesBookingResponse CalculateFeesSeatAssignment(CalculateFeesSeatAssignmentRequest Request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            CalculateFeesBookingResponse response = new CalculateFeesBookingResponse();

            Booking booking = new Booking();

            // map from message to entity
            if (Request != null)
            {
                booking.Header = Request.Booking.Header.ToEntity();
                booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                booking.Passengers = Request.Booking.Passengers.ToListEntity();
                booking.Mappings = Request.Booking.Mappings.ToListEntity();
                booking.Payments = Request.Booking.Payments.ToListEntity();
                booking.Fees = Request.Booking.Fees.ToListEntity();
                booking.Remarks = Request.Booking.Remarks.ToListEntity();
                booking.Services = Request.Booking.Services.ToListEntity();
                booking.Taxs = Request.Booking.Taxs.ToListEntity();
            }

            try
            {
                IList<Avantik.Web.Service.Entity.Booking.Fee> BookingFees = objFeeService.CalculateFeesSeatAssignment(Request.AgencyCode, Request.Currency, booking, Request.Language, Request.bNovat);

                if (BookingFees != null)
                {
                    // map to response
                    response.Fees = BookingFees.ToBookingMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "CalculateFeesSeatAssignment error";
                response.Success = false;
            }

            return response;
        }
        
        public CalculateFeesBookingResponse CalculateFeesNameChange(CalculateFeesNameChangeRequest Request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            CalculateFeesBookingResponse response = new CalculateFeesBookingResponse();

            Booking booking = new Booking();

            // map from message to entity
            if (Request != null)
            {
                booking.Header = Request.Booking.Header.ToEntity();
                booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                booking.Passengers = Request.Booking.Passengers.ToListEntity();
                booking.Mappings = Request.Booking.Mappings.ToListEntity();
                booking.Payments = Request.Booking.Payments.ToListEntity();
                booking.Fees = Request.Booking.Fees.ToListEntity();
                booking.Remarks = Request.Booking.Remarks.ToListEntity();
                booking.Services = Request.Booking.Services.ToListEntity();
                booking.Taxs = Request.Booking.Taxs.ToListEntity();
            }

            try
            {
              //  IList<Avantik.Web.Service.Entity.Booking.Fee> BookingFees = objFeeService.CalculateFeesNameChange(Request.AgencyCode, Request.Currency, booking, Request.Language);

                //if (BookingFees != null)
                //{
                //    // map to response
                //    response.Fees = BookingFees.ToBookingMessage();
                //    response.Message = "Success";
                //    response.Success = true;
                //}
                //else
                //{
                //    response.Message = "Fail";
                //    response.Success = false;
                //}
            }
            catch
            {
                response.Message = "CalculateFeesNameChange error";
                response.Success = false;
            }

            return response;
        }
        
        public CalculateFeesBookingResponse CalculateFeesSpecialService(CalculateFeesSpecialServiceRequest Request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            CalculateFeesBookingResponse response = new CalculateFeesBookingResponse();

            Booking booking = new Booking();

            // map from message to entity
            if (Request != null)
            {
                booking.Header = Request.Booking.Header.ToEntity();
                booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                booking.Passengers = Request.Booking.Passengers.ToListEntity();
                booking.Mappings = Request.Booking.Mappings.ToListEntity();
                booking.Payments = Request.Booking.Payments.ToListEntity();
                booking.Fees = Request.Booking.Fees.ToListEntity();
                booking.Remarks = Request.Booking.Remarks.ToListEntity();
                booking.Services = Request.Booking.Services.ToListEntity();
                booking.Taxs = Request.Booking.Taxs.ToListEntity();
            }

            try
            {
                Avantik.Web.Service.Entity.Booking.Booking BookingFees = objFeeService.CalculateFeesSpecialService(Request.AgencyCode, Request.Currency, booking, Request.Language, Request.bNovat);

                if (BookingFees != null)
                {
                    // map to response
                    //response.Fees = BookingFees;
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }

            }
            catch
            {
                response.Message = "CalculateFeesSpecialService error";
                response.Success = false;
            }

            return response;
        }
        
        public GetSegmentFeeResponse GetServiceFee(GetSegmentFeeRequest Request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            GetSegmentFeeResponse response = new GetSegmentFeeResponse();

            IList<Entity.Booking.PassengerService> passengerServices = Request.services.ToListEntity();
            IList<Entity.SegmentService> segmentServices = Request.segmentService.ToFeeEntity();

            try
            {
                IList<Entity.ServiceFee> ServiceFees = objFeeService.GetSegmentFee(Request.AgencyCode, Request.CurrencyCode, Request.LanguageCode, Request.NumberOfPassenger, Request.NumberOfInfant, passengerServices, segmentServices, Request.SpecialService, Request.bNoVat);

                if (ServiceFees != null)
                {
                    // map to response
                    response.ServiceFees = ServiceFees.ToFeeMessage();
                    response.Message = "Success";
                    response.Success = true;
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "GetSegmentFee error";
                response.Success = false;
            }

            return response;
        }

        public BaggageFeeResponse GetBaggageFee(BaggageFeeRequest request)
        {
            IFeeService objFeeService = FeeServiceFactory.CreateInstance();
            BaggageFeeResponse response = new BaggageFeeResponse();

            try
            {
                if (string.IsNullOrEmpty(request.AgencyCode))
                {
                    response.Success = false;
                    response.Message = "Agency Code is required";
                }
                else if (request.BookingSegmentId.Equals(Guid.Empty))
                {
                    response.Success = false;
                    response.Message = "Booking Segment Id is required";
                }
                else if (request.Mappings.Count == 0)
                {
                    response.Success = false;
                    response.Message = "Mappings is required";
                }
                else
                {
                    if (string.IsNullOrEmpty(request.LanguageCode))
                    {
                        request.LanguageCode = "EN";
                    }

                    IList<Entity.Booking.Fee> bkf = request.BookingFees.ToListEntity();
                    IList<Entity.Booking.Mapping> mpp = request.Mappings.ToListEntity();

                    IList<Avantik.Web.Service.Entity.Booking.Fee> fee = objFeeService.GetBaggageFee(mpp, 
                                                                 request.BookingSegmentId, 
                                                                 request.PassengerId, 
                                                                 request.AgencyCode, 
                                                                 request.LanguageCode, 
                                                                 request.MaxUnit,
                                                                 bkf, 
                                                                 request.bNovat);
                    if (fee != null && fee.Count > 0)
                    {
                        response.Success = true;
                        response.Message = "Success";
                    }
                    else
                    {
                        response.Success = false;
                        response.Message = "Failed";
                    }
                }
            }
            catch (SystemException ex)
            {
                response.Success = false;
                response.Message = ex.Message;

                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(BaggageFeeRequest), request));
            }
            return response;
        }

        public QuoteSummaryResponse GetQuoteSummary(QuoteSummaryRequest Request)
        {
            IBookingModelService objBookingService = BookingServiceFactory.CreateInstance();
            QuoteSummaryResponse response = new QuoteSummaryResponse();


            try
            {
                if (Request.FlightSegments != null && Request.FlightSegments.Count > 0)
                {
                    IList<Entity.Booking.FlightSegment> flightSegments = Request.FlightSegments.ToListEntity();
                    IList<Entity.Booking.Passenger> passengers = Request.Passengers.ToListEntity();
                    Booking booking = objBookingService.GetQuoteSummary(flightSegments, passengers, Request.AgencyCode, Request.Language, Request.CurrencyCode, Request.bNovat);

                    if (booking != null && booking.Segments != null && booking.Segments.Count > 0)
                    {
                        response.BookingResponse.FlightSegments = booking.Segments.ToBookingMessage();
                        response.BookingResponse.Passengers = booking.Passengers.ToBookingMessage();
                        response.BookingResponse.Taxs = booking.Taxs.ToBookingMessage();
                        response.BookingResponse.Mappings = booking.Mappings.ToBookingMessage();
                        response.BookingResponse.Fees = booking.Fees.ToBookingMessage();
                        response.BookingResponse.Quotes = booking.Quotes.ToBookingMessage();

                        response.Message = "Success";
                        response.Success = true;
                    }
                    else
                    {
                        response.Message = "Fail";
                        response.Success = false;
                    }
                }
                else
                {
                    response.Message = "Fail";
                    response.Success = false;
                }
            }
            catch(SystemException ex)
            {
                response.Message = "Booking Read error";
                response.Success = false;
                //Logger.Instance(Logger.LogType.Mail).WriteLog(ex, XMLHelper.JsonSerializer(typeof(QuoteSummaryRequest), Request));
            }

            return response;
        }

        public PaymentMultipleFOPResponse PaymentMultipleFOP(PaymentMultipleFOPRequest Request)
        {
            IAgencyService objAgencyService = AgencyServiceFactory.CreateInstance();
            IPaymentService objPaymentService = PaymentServiceFactory.CreateInstance();
            IVoucherService objVoucherService = VoucherServiceFactory.CreateInstance();
            PaymentMultipleFOPResponse response = new PaymentMultipleFOPResponse();

            Booking booking = null;
            bool resultPayment = false;
            IList<PaymentAllocation> paymentAllocations = null;

            List<Entity.Booking.Payment> PaymentsVoucher = new List<Entity.Booking.Payment>();
            List<Entity.Booking.Payment> PaymentsCreitCard = new List<Entity.Booking.Payment>();
            List<Entity.Booking.Payment> PaymentsCreditAgency = new List<Entity.Booking.Payment>();
            List<Entity.Booking.Payment> PaymentsToSave = new List<Entity.Booking.Payment>();
            Entity.Voucher voucher = new Entity.Voucher();
            PaymentsToSave = Request.Booking.Payments.ToListEntity() as List<Entity.Booking.Payment>;
            Comparison<Entity.Booking.Payment> comparePaymentAmount = new Comparison<Entity.Booking.Payment>(Entity.Booking.Payment.ComparePaymentAmount);

            try
            {
                if (Request != null)
                {
                    booking = new Booking();
                    booking.Header = Request.Booking.Header.ToEntity();
                    booking.Segments = Request.Booking.FlightSegments.ToListEntity();
                    booking.Passengers = Request.Booking.Passengers.ToListEntity();
                    booking.Remarks = Request.Booking.Remarks.ToListEntity();
                    booking.Payments = Request.Booking.Payments.ToListEntity();
                    booking.Mappings = Request.Booking.Mappings.ToListEntity();
                    booking.Services = Request.Booking.Services.ToListEntity();
                    booking.Taxs = Request.Booking.Taxs.ToListEntity();
                    booking.Fees = Request.Booking.Fees.ToListEntity();

                    foreach (Entity.Booking.Payment p in booking.Payments)
                    {
                        if (p.FormOfPaymentRcd.Equals("CC"))
                        {
                            PaymentsCreitCard.Add(p);
                        }
                        else if (p.FormOfPaymentRcd.Equals("VOUCHER"))
                        {
                            voucher.VoucherNumber = p.DocumentNumber;
                         //   voucher = objVoucherService.GetVoucher(voucher);

                            if (voucher == null)
                            {
                                resultPayment = false;
                            }
                            else
                            {
                                PaymentsVoucher.Add(p);
                            }

                        }
                        else if (p.FormOfPaymentRcd.Equals("INV") || p.FormOfPaymentRcd.Equals("CRAGT"))
                        {
                            PaymentsCreditAgency.Add(p);
                        }
                    }

                    if (booking.ValidateSaveBooking(booking.Header, booking.Segments, booking.Passengers, booking.Mappings))
                    {
                        if (PaymentsVoucher.Count > 0)
                        {
                            PaymentsVoucher.Sort(comparePaymentAmount);
                            PaymentsToSave.AddRange(PaymentsVoucher);
                            if (voucher.ValidateVoucherPayment(PaymentsVoucher) == true)
                            {
                                if (PaymentsCreitCard.Count > 0)
                                {
                                    PaymentsToSave.AddRange(PaymentsCreitCard);
                                    paymentAllocations = booking.CreateAllocation();
                                    bool validCreditCard = false;

                                    // validCreditCard = objPaymentService.PaymentCreditCard(booking, paymentAllocations,
                                    //                                                    Request.SecurityToken, Request.AuthenticationToken,
                                    //                                                    Request.CommerceIndicator, Request.BookingReference,
                                    //                                                    Request.RequestSource, true, false, false);
                                    if (validCreditCard)
                                    {
                                        resultPayment = objPaymentService.SavePayment(PaymentsToSave, paymentAllocations);

                                        if (resultPayment)
                                        {
                                            response.Message = "Success";
                                            response.Success = true;
                                        }
                                        else
                                        {
                                            response.Message = "SavePayment CC && Voucher failed.";
                                            response.Success = false;
                                        }
                                    }
                                }
                                if (PaymentsCreditAgency.Count > 0)
                                {
                                    Agent agent = new Agent();
                                    bool agentResult = true;
                                    PaymentsToSave.AddRange(PaymentsCreditAgency);
                                    paymentAllocations = booking.CreateAllocation();

                                    agent = objAgencyService.GetAgencySessionProfile(Request.Booking.Header.AgencyCode, string.Empty);
                                    if (agent != null)
                                    {

                                        foreach (Entity.Booking.Payment p in PaymentsCreditAgency)
                                        {
                                            if (!p.ValidateCreditAgency(agent))
                                            {
                                                agentResult = false;
                                                response.Message = "ValidateCreditAgency fail";
                                                response.Success = false;
                                            }
                                        }

                                        if (agentResult)
                                        {
                                            booking = new Booking();
                                            booking.Payments = Request.Booking.Payments.ToListEntity();
                                            booking.Mappings = Request.Booking.Mappings.ToListEntity();
                                            booking.Fees = Request.Booking.Fees.ToListEntity();
                                            paymentAllocations = booking.CreateAllocation();
                                            resultPayment = objPaymentService.SavePayment(PaymentsToSave, paymentAllocations);

                                            if (resultPayment)
                                            {
                                                response.Message = "Success";
                                                response.Success = true;
                                            }
                                            else
                                            {
                                                response.Message = "SavePayment Agent && Voucher fail";
                                                response.Success = false;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        response.Message = "GetAgencySessionProfile fail";
                                        response.Success = false;
                                    }

                                    
                                }
                            }
                        }
                    }

                    if (resultPayment == true)
                    {
                        response.Message = "Success";
                        response.Success = true;
                    }
                    else
                    {
                        response.Message = "Multiple Payment fail";
                        response.Success = false;
                    }
                }
                else
                {
                    response.Message = "Booking Multiple Payment is null";
                    response.Success = false;
                }
            }
            catch
            {
                response.Message = "Booking Multiple Payment Error";
                response.Success = false;
            }

            return response;
        }

        public GetSeatMapResponse GetSeatMap(GetSeatMapRequest request)
        {
            Model.Contract.IFlightService objFlightService = FlightServiceFactory.CreateInstance();
            GetSeatMapResponse response = new GetSeatMapResponse();
            List<Entity.Flight.SeatMap> seatMaps = new List<Entity.Flight.SeatMap>();
            Entity.Booking.Flight flight = new Entity.Booking.Flight();

            // map request to entity
            flight.OriginRcd = request.OriginRcd;
            flight.DestinationRcd = request.DestinationRcd;
            flight.FlightId = new Guid(request.FlightId);
            flight.BoardingClassRcd = request.BoardingClass;
            flight.BookingClassRcd = request.BookingClass;

            try
            {
                // because seat is booked always can not to cache
               // seatMaps = (List<Entity.Flight.SeatMap>)HttpRuntime.Cache["SeatMap-" + flight.FlightId.ToString().ToUpper() + flight.BoardingClassRcd.ToUpper()];

                if (seatMaps == null || seatMaps.Count == 0)
                {
                    seatMaps = objFlightService.GetSeatMap(flight);

                  //  HttpRuntime.Cache.Insert("SeatMap-" + flight.FlightId.ToString().ToUpper() + flight.BoardingClassRcd.ToUpper(), seatMaps, null, System.Web.Caching.Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(20), System.Web.Caching.CacheItemPriority.Normal, null);
                }

                if (seatMaps != null && seatMaps.Count > 0)
                {
                    response.SeatMaps = seatMaps.ToSeatMapMessage();
                    response.Success = true;
                    response.Message = "Success";
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
    }
}
