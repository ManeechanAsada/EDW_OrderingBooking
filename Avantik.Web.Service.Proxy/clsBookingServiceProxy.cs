
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Contracts;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Fee;
using Avantik.Web.Service.Message.SeatMap;

namespace Avantik.Web.Service.Proxy
{
    public class BookingServiceProxy : ClientBase<IBookingService>, IBookingService
    {
       // public BookingSaveResponse SaveBooking(BookingSaveRequest request)
       // {
          //  return base.Channel.SaveBooking(request);
       // }
        //public BookingReadResponse ReadBooking(BookingReadRequest request)
        //{
        //    return base.Channel.ReadBooking(request);
        //}

      //  public BookingPaymentResponse BookingPayment(BookingPaymentRequest request)
       // {
       //     return base.Channel.BookingPayment(request);
       // }
        //public AgencySessionProfileResponse GetAgencySessionProfile(AgencySessionProfileRequest request)
        //{
        //    return base.Channel.GetAgencySessionProfile(request);
        //}
        //public TravelAgentLogonResponse TravelAgentLogon(TravelAgentLogonRequest request)
        //{
        //    return base.Channel.TravelAgentLogon(request);
        //}
        //public BookingFlightResponse BookFlight(BookingFlightRequest request)
        //{
        //    return base.Channel.BookFlight(request);
        //}
        //public BookingPaymentCreditCardResponse BookingPaymentCreditCard(BookingPaymentCreditCardRequest request)
        //{
        //    return base.Channel.BookingPaymentCreditCard(request);
        //}
        //public BookingReadVoucherResponse BookingReadVoucher(BookingReadVoucherRequest request)
        //{
        //    return base.Channel.BookingReadVoucher(request);
        //}
        //public BookingPaymentVoucherResponse BookingPaymentVoucher(BookingPaymentVoucherRequest request)
        //{
        //    return base.Channel.BookingPaymentVoucher(request);
        //}
        //public BookingPaymentCreditAgencyResponse BookingPaymentCreditAgency(BookingPaymentCreditAgencyRequest request)
        //{
        //    return base.Channel.BookingPaymentCreditAgency(request);
        //}

        //public  GetFeeResponse GetFee(GetFeeRequest request)
        //{
        //    return base.Channel.GetFee(request);
        //}
        //public CalculateFeesBookingResponse CalculateFeesBookingCreate(CalculateFeesBookingCreateRequest request)
        //{
        //    return base.Channel.CalculateFeesBookingCreate(request);
        //}
        //public CalculateFeesBookingResponse CalculateFeesBookingChange(CalculateFeesBookingCreateRequest request)
        //{
        //    return base.Channel.CalculateFeesBookingChange(request);
        //}
        //public CalculateFeesBookingResponse CalculateFeesSeatAssignment(CalculateFeesSeatAssignmentRequest request)
        //{
        //    return base.Channel.CalculateFeesSeatAssignment(request);
        //}
        //public CalculateFeesBookingResponse CalculateFeesNameChange(CalculateFeesNameChangeRequest request)
        //{
        //    return base.Channel.CalculateFeesNameChange(request);
        //}
        //public CalculateFeesBookingResponse CalculateFeesSpecialService(CalculateFeesSpecialServiceRequest request)
        //{
        //    return base.Channel.CalculateFeesSpecialService(request);
        //}
        //public GetSegmentFeeResponse GetServiceFee(GetSegmentFeeRequest request)
        //{
        //    return base.Channel.GetServiceFee(request);
        //}
        //public BaggageFeeResponse GetBaggageFee(BaggageFeeRequest request)
        //{
        //    return base.Channel.GetBaggageFee(request);
        //}

        //public QuoteSummaryResponse GetQuoteSummary(QuoteSummaryRequest request)
        //{
        //    return base.Channel.GetQuoteSummary(request);
        //}

        //public PaymentMultipleFOPResponse PaymentMultipleFOP(PaymentMultipleFOPRequest request)
        //{
        //    return base.Channel.PaymentMultipleFOP(request);
        //}
        //public GetSeatMapResponse GetSeatMap(GetSeatMapRequest request)
        //{
        //    return base.Channel.GetSeatMap(request);
        //}
    }
}
