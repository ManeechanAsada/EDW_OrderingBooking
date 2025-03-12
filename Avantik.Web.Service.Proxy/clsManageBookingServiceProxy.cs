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
using Avantik.Web.Service.Message.ManageBooking;
using Avantik.Web.Service.Message.SeatMap;
using Avantik.Web.Service.Message.OrderBooking;

namespace Avantik.Web.Service.Proxy
{
    public class ManageBookingServiceProxy : ClientBase<IManageBookingService>, IManageBookingService
    {
        public InitializeResponse ServiceInitialize(InitializeRequest request)
        {
            return base.Channel.ServiceInitialize(request);
        }
        //public AvailabilityResponse GetAvailability(AvailabilityRequest request)
        //{
        //    return base.Channel.GetAvailability(request);
        //}
        public BookingCancelResponse CancelBooking(BookingCancelRequest request)
        {
            return base.Channel.CancelBooking(request);
        }
        public BookingReadResponse ReadBooking(BookingReadRequest request)
        {
            return base.Channel.ReadBooking(request);
        }
        public SegmentCancelResponse CancelSegment(SegmentCancelRequest request)
        {
            return base.Channel.CancelSegment(request);
        }
        public ModifyBookingResponse ModifyBooking(ModifyBookingRequest request)
        {
            return base.Channel.ModifyBooking(request);
        }
        public ModifyBookingResponse OrderBooking(OrderBookingRequest request)
        {
            return base.Channel.OrderBooking(request);
        }

        public GetSeatMapResponse GetSeatMap(GetSeatMapRequest request)
        {
            return base.Channel.GetSeatMap(request);
        }
        //public CalculateChangeFlightFeesResponse GetChangeFlightFees(CalculateChangeFlightFeesRequest request)
        //{
        //    return base.Channel.GetChangeFlightFees(request);
        //}
        //public CalculateSeatFeesResponse SelectSeat(CalculateSeatFeesRequest request)
        //{
        //    return base.Channel.SelectSeat(request);
        //}
        public GetSpecialServicesResponse GetSpecialServices(GetSpecialServicesRequest request)
        {
            return base.Channel.GetSpecialServices(request);
        }
        //public CalculateSpecialServiceFeesResponse AddSpecialServices(CalculateSpecialServiceFeesRequest request)
        //{
        //    return base.Channel.AddSpecialServices(request);
        //}
        //public BaggageFeesResponse GetBaggageFee(BaggageFeesRequest request)
        //{
        //    return base.Channel.GetBaggageFee(request);
        //}
        //public PassengerInfoResponse UpdatePassengerInfo(PassengerInfoRequest request)
        //{
        //    return base.Channel.UpdatePassengerInfo(request);
        //}
        //public UpdatedTicketResponse UpdateTicketBaggageAllowance(UpdatedTicketRequest request)
        //{
        //    return base.Channel.UpdateTicketBaggageAllowance(request);
        //}
        //public NameChangeResponse GetNameChangeFee(NameChangeRequest request)
        //{
        //    return base.Channel.GetNameChangeFee(request);
        //}
        public GetFeeResponse GetFeeDefinition(GetFeeDefinitionRequest request)
        {
            return base.Channel.GetFeeDefinition(request);
        }
        //  public PaymentFeeResponse GetPaymentFee(PaymentFeeRequest request)
        //  {
        //     return base.Channel.GetPaymentFee(request);
        // }
        //public UpdatedTicketResponse UpdateTicketBaggageAllowance(UpdatedTicketRequest request)
        //{
        //    return base.Channel.UpdateTicketBaggageAllowance(request);
        //}
        //public AgencyResponse GetAgencyProfile(AgencyRequest request)
        //{
        //    return base.Channel.GetAgencyProfile(request);
        //}
        //public ContactDetailResponse UpdateContactDetail(ContactDetailRequest request)
        //{
        //    return base.Channel.UpdateContactDetail(request);
        //}


    }
}
