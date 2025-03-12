using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Message.Fee;
using Avantik.Web.Service.Message.SeatMap;

namespace Avantik.Web.Service.Contracts
{
    [ServiceContract(Namespace = "Avantik.Web.Service/")]
    public interface IBookingService 
    {
        
       // [OperationContract()]
      //  BookingSaveResponse SaveBooking(BookingSaveRequest request);

       // [OperationContract()]
      //  BookingReadResponse ReadBooking(BookingReadRequest request);

        //[OperationContract()]
        //AgencySessionProfileResponse GetAgencySessionProfile(AgencySessionProfileRequest request);

        //[OperationContract()]
        //TravelAgentLogonResponse TravelAgentLogon(TravelAgentLogonRequest request);

      //  [OperationContract()]
      //  BookingPaymentResponse BookingPayment(BookingPaymentRequest request);

      //  [OperationContract()]
     //   BookingFlightResponse BookFlight(BookingFlightRequest request);

        //[OperationContract()]
        //BookingPaymentCreditCardResponse BookingPaymentCreditCard(BookingPaymentCreditCardRequest request);

        //[OperationContract()]
        //BookingReadVoucherResponse BookingReadVoucher(BookingReadVoucherRequest request);

        //[OperationContract()]
        //BookingPaymentVoucherResponse BookingPaymentVoucher(BookingPaymentVoucherRequest request);

        //[OperationContract()]
        //BookingPaymentCreditAgencyResponse BookingPaymentCreditAgency(BookingPaymentCreditAgencyRequest request);
        
        //[OperationContract()]
        //GetFeeResponse GetFee(GetFeeRequest Request);

        //[OperationContract()]
        //CalculateFeesBookingResponse CalculateFeesBookingCreate(CalculateFeesBookingCreateRequest Request);

        //[OperationContract()]
        //CalculateFeesBookingResponse CalculateFeesBookingChange(CalculateFeesBookingCreateRequest Request);

        //[OperationContract()]
        //CalculateFeesBookingResponse CalculateFeesSeatAssignment(CalculateFeesSeatAssignmentRequest Request);

        //[OperationContract()]
        //CalculateFeesBookingResponse CalculateFeesNameChange(CalculateFeesNameChangeRequest Request);

        //[OperationContract()]
        //CalculateFeesBookingResponse CalculateFeesSpecialService(CalculateFeesSpecialServiceRequest Request);

      //  [OperationContract()]
      //  GetSegmentFeeResponse GetServiceFee(GetSegmentFeeRequest Request);

        //[OperationContract()]
        //BaggageFeeResponse GetBaggageFee(BaggageFeeRequest Request);

        //[OperationContract()]
        //QuoteSummaryResponse GetQuoteSummary(QuoteSummaryRequest request);

        //[OperationContract()]
        //PaymentMultipleFOPResponse PaymentMultipleFOP(PaymentMultipleFOPRequest request);

     //   [OperationContract()]
      //  GetSeatMapResponse GetSeatMap(GetSeatMapRequest request);

    }
    
}
