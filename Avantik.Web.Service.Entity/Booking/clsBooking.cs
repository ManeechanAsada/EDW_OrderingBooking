using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Agency;
using Avantik.Web.Service.Entity;

namespace Avantik.Web.Service.Entity.Booking
{
    public class Booking
    {
        public BookingHeader Header { get; set; }
        public IList<FlightSegment> Segments { get; set; }
        public IList<Passenger> Passengers { get; set; }
        public IList<Remark> Remarks { get; set; }
        public IList<Payment> Payments { get; set; }
        public IList<Mapping> Mappings { get; set; }
        public IList<PassengerService> Services { get; set; }
        public IList<Tax> Taxs { get; set; }
        public IList<Fee> Fees { get; set; }
        public IList<Quote> Quotes { get; set; }

        #region Method
        // create allocate payment
        public IList<PaymentAllocation> CreateAllocation()
        {
            IList<PaymentAllocation> paymentAllocations = new List<PaymentAllocation>();


            if (Payments != null && Mappings != null)
            {
                foreach (Avantik.Web.Service.Entity.Booking.Payment pay in Payments)
                {
                    foreach (Avantik.Web.Service.Entity.Booking.Mapping m in Mappings)
                    {
                        if (m != null)
                        {
                            foreach(Avantik.Web.Service.Entity.Booking.Mapping mp in Mappings)
                            {
                                if ((!(mp.ExchangedDateTime == DateTime.MinValue) && (m.BookingSegmentId.ToString() == mp.ExchangedSegmentId.ToString()) && (m.PassengerId.ToString() == mp.PassengerId.ToString())))
                                {
                                    m.ExchangedPaid = mp.PaymentAmount;
                                    break;
                                }
                            }


                            // find new charge
                            //if (m.TransferableFareFlag == 1 && m.PaymentAmount == 0 && x > 1)
                            //{
                            //    for (int j = 0; j < Mappings.Count; j++)
                            //    {
                            //        if (m.PassengerId == Mappings[j].PassengerId
                            //            && m.PassengerStatusRcd.ToUpper() == "OK"
                            //            && Mappings[j].PassengerStatusRcd.ToUpper() == "XX")
                            //        {
                            //            if (m.FareAmountIncl == Mappings[j].FareAmountIncl)
                            //            {
                            //                transfer = Mappings[j].FareAmountIncl;
                            //                break;
                            //            }
                            //            else if (m.FareAmountIncl < Mappings[j].FareAmountIncl)
                            //            {
                            //                transfer = Mappings[j].FareAmountIncl;
                            //                break;
                            //            }
                            //            else if (m.FareAmountIncl > Mappings[j].FareAmountIncl)
                            //            {
                            //                transfer = m.PaymentAmount - Mappings[j].PaymentAmount;
                            //                break;
                            //            }
                            //        }
                            //    }
                            //}


                            PaymentAllocation allocation = m.GetAllocation(pay.BookingPaymentId, pay.CreateBy);

                            if (allocation != null)
                                paymentAllocations.Add(allocation);
                        }
                    }

                    if (Fees != null)
                    {
                        foreach (Avantik.Web.Service.Entity.Booking.Fee f in Fees)
                        {
                            if (f != null && f.PaymentAmount == 0 && f.VoidDateTime == DateTime.MinValue)
                           
                            {
                                PaymentAllocation allocation = f.GetAllocation(pay.BookingPaymentId, pay.CreateBy);

                                if (allocation != null)
                                    paymentAllocations.Add(allocation);
                            }
                        }
                    }

                }
            }

            return paymentAllocations;
        }
        public bool FindUSSegment()
        {
            try
            {
                if (Segments.Count > 0)
                {
                    foreach (FlightSegment se in Segments)
                    {
                        if (se.SegmentStatusRcd == "US")
                        {
                            return true;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }

            return false;
        }
        public FlightSegment GetUSSegment()
        {
            try
            {
                if (Segments.Count > 0)
                {
                    foreach (FlightSegment se in Segments)
                    {
                        if (se.SegmentStatusRcd == "US")
                        {
                            return se;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }

            return null;
        }

        public bool ValidateSaveBooking(BookingHeader headers, IList<FlightSegment> segments, IList<Passenger> passengers, IList<Mapping> mappings)
        {
            bool result = false;

            if(headers == null)
                result = false;
            else if (segments == null || segments.Count == 0)
                result = false;
            else if(passengers == null || passengers.Count == 0)
                result = false;
            else if(mappings == null || mappings.Count == 0)
                result = false;
            else{
                if(headers.GroupName.Equals(string.Empty)){
                    for(int i = 0; i < passengers.Count ; i++){
                        if(passengers[i].Lastname.Equals(string.Empty) || passengers[i].Lastname.Length == 0){
                            result = false;
                            break;
                        }
                    }

                    for(int i = 0; i < mappings.Count ; i++){
                        if(mappings[i].Lastname.Equals(string.Empty) || mappings[i].Lastname.Length == 0){
                            result = false;
                            break;
                        }
                    }
                }
            }
            return result;
        }

        public bool ValidateVoucherDuplicate(IList<Voucher> vouchers)
        {
            return ValidateVoucherDuplicate(vouchers);
        }
        public bool ValidateVoucherPaymentAmount(IList<Voucher> vouchers, Payment payment)
        {
            return ValidateVoucherPaymentAmount(vouchers, payment);
        }
        public bool ValidateCreditAgencyFOP()
        {
            return ValidateCreditAgencyFOP();
        }
        public bool ValidateCreditAgencyCurrencyRCD(Agent agent)
        {
            return ValidateCreditAgencyCurrencyRCD(agent);
        }
        public bool ValidateCreditAgencyBalance(Agent agent)
        {
            return ValidateCreditAgencyBalance(agent);
        }

        public void SetBookingSource(byte ownAgencyFlag, byte webAgencyFlag)
        {
            if (Header != null)
            {
                if (ownAgencyFlag == 1 && webAgencyFlag == 1)
                    Header.BookingSourceRcd = "B2C";
                else if (ownAgencyFlag == 0 && webAgencyFlag == 1)
                    Header.BookingSourceRcd = "B2B";
            }
        }

        public decimal CalOutStandingBalance(IList<Fee> fees, IList<Mapping> mappings)
        {
            decimal dclTotalTicket = 0;
            decimal dclTotalPayment = 0;
            decimal TotalOutStanding = 0;
            decimal dclChargeAmount = 0;
            decimal dclPaymentAmount = 0;

            //FindTotal Ticket
            for (int i = 0; i < mappings.Count; i++)
            {
                dclChargeAmount = mappings[i].NetTotal;

                if (mappings[i].RefundDateTime.Equals(DateTime.MinValue) == false)
                {
                    dclChargeAmount = mappings[i].RefundCharge;
                }
                else if (mappings[i].ExcludePricingFlag != 0)
                {
                    dclChargeAmount = 0;
                }

                dclTotalTicket = dclTotalTicket + dclChargeAmount;

                if (!string.IsNullOrEmpty(mappings[i].PassengerStatusRcd))
                {
                    if ((mappings[i].PassengerStatusRcd.ToUpper() == "XX" | mappings[i].PassengerStatusRcd.ToUpper() == "XK") & string.IsNullOrEmpty(mappings[i].TicketNumber))
                    {
                        dclPaymentAmount = 0;
                    }
                    else
                    {
                        dclPaymentAmount = mappings[i].PaymentAmount;
                    }
                }

                dclTotalPayment = dclTotalPayment + dclPaymentAmount;
            }
            
            // fee
            if (fees != null)
            {
                foreach (Fee f in fees)
                {
                    if (f.VoidDateTime == DateTime.MinValue)
                    {
                        dclTotalTicket = dclTotalTicket + f.FeeAmountIncl;
                        dclTotalPayment = dclTotalPayment + f.PaymentAmount;
                    }
                }
            }

            TotalOutStanding = dclTotalTicket - dclTotalPayment;

            return TotalOutStanding;
        }
        
        public Boolean IsClosedForCheckIn()
        {
            Boolean result = false;
            if (!string.IsNullOrEmpty(Segments[0].FlightCheckInStatusRcd) && Segments[0].FlightCheckInStatusRcd.Equals("CLOSED"))
                result = true;
            return result;
        }
        
        public Boolean IsOpenForCheckIn()
        {
            Boolean result = false;
            if (!string.IsNullOrEmpty(Segments[0].FlightCheckInStatusRcd) && Segments[0].FlightCheckInStatusRcd.Equals("OPEN"))
                result = true;
            return result;
        }
       
        public Boolean CheckFlightCheckInStatusRcd()
        {
            // may be wrong
            Boolean result = false;
            if (string.IsNullOrEmpty(Segments[0].FlightCheckInStatusRcd ))
                result = true;
            return result;
        }

        public Boolean CheckPassengerStatus(Int32 index)
        {
            Boolean result = false;
            List<Mapping> mappings = Mappings.Where(item => item.BookingSegmentId == Segments[index].BookingSegmentId).ToList();

            foreach (Mapping mp in mappings)
            {
                if (mp.PassengerStatusRcd.Equals("OK") && mp.BookingSegmentId.Equals(Segments[index].BookingSegmentId))
                {
                    result = true;
                }
            }
            return result;
        }
        
        public Boolean CheckPassengerStatusInBooking()
        {
            Boolean result = false;
            Int32 total = 0;
            total = Mappings.Count(item => item.PassengerStatusRcd != "OK" && item.PassengerStatusRcd != "XX");

            if (total == 0)
                result = true;

            return result;
        }

        public Boolean CheckPassengerCheckInInBooking()
        {
            Boolean result = false;
            Int32 total = 0;
            total = Mappings.Count(item => item.CheckInBy != Guid.Empty);

            if (total == 0)
                result = true;

            return result;
        }

        // validate booking
        public Boolean IsLockBooking()
        {
            Boolean result = false;

            if (!string.IsNullOrEmpty(Header.LockName) && !Header.LockName.Trim().Equals(""))
            {
                result = true;
            }
               
            return result;
        }
        
        //Seat validation
        public Boolean IsDuplicateSeat(String seatNumber)
        {
            Boolean result = false;
           // Int32 total = 0;
            // Int32 total = Mappings.Count(item => item.SeatNumber.Equals(seatNumber));

            foreach (Mapping m in this.Mappings)
            {
                if (!string.IsNullOrEmpty(m.SeatNumber) && m.SeatNumber.Trim().ToUpper() == seatNumber.Trim().ToUpper())
                {
                    result = true;
                    break;
                }

            }
            //if (total > 0)
            //    result = true;

            return result;
        }
        public Boolean IsAlreadyAssignSeat(String segmentId, String passengerId)
        {
            Boolean result = false;
            try
            {
                Int32 total = Mappings.Count(item => item.BookingSegmentId == Guid.Parse(segmentId) && item.PassengerId == Guid.Parse(passengerId) && item.SeatNumber.Equals(string.Empty));

                if (total > 0)
                    result = true;
            }
            catch
            {
                throw;
            }

            return result;
        }
        public Boolean IsFlightStatusAllowAssignSeat(String segmentId)
        {
            Boolean result = false;
            try
            {
                Int32 total = Segments.Count(item => item.BookingSegmentId == Guid.Parse(segmentId) && (item.FlightCheckInStatusRcd != null) && (item.FlightCheckInStatusRcd.Trim().ToUpper().Equals("OPEN") || item.FlightCheckInStatusRcd.ToUpper().Equals(String.Empty)));

                if (total > 0)
                    result = true;
            }
            catch
            {
                throw;
            }
            return result;
        }
      
        // validate segment
        public Boolean IsSegmentInBooking(String segmentId)
        {
            Boolean result = false;
            try
            {
                Int32 total = Mappings.Count(item => item.BookingSegmentId == Guid.Parse(segmentId));

                if (total > 0)
                    result = true;
            }
            catch
            {
                throw;
            }
            return result;
        }
        public Boolean IsValidSegmentStatus(String segmentId)
        {
            Boolean result = false;
            foreach (Mapping mp in Mappings)
            {
                if (mp.BookingSegmentId == new Guid(segmentId))
                {
                    if (!string.IsNullOrEmpty(mp.PassengerStatusRcd) && mp.PassengerStatusRcd.Trim().ToUpper().Equals("OK"))
                    {
                        result = true;
                        break;
                    }
                }
            }

            if (result)
            {
                foreach (FlightSegment fs in Segments)
                {
                    if (fs.BookingSegmentId == new Guid(segmentId))
                    {
                        if (!string.IsNullOrEmpty(fs.SegmentStatusRcd) && fs.SegmentStatusRcd.Trim().ToUpper().Equals("HK"))
                        {
                            result = true;
                            break;
                        }
                    }
                }
            }

            return result;
        }

        // check all booking
        public Boolean IsValidSegmentStatus()
        {
            Boolean result = false;
            foreach (Mapping mp in Mappings)
            {
                if (!string.IsNullOrEmpty(mp.PassengerStatusRcd) && mp.PassengerStatusRcd.Trim().ToUpper().Equals("OK"))
                {
                    result = true;
                    break;
                }
            }

            if (result)
            {
                foreach (FlightSegment fs in Segments)
                {
                    if (!string.IsNullOrEmpty(fs.SegmentStatusRcd) && fs.SegmentStatusRcd.Trim().ToUpper().Equals("HK"))
                    {
                        result = true;
                        break;
                    }
                }
            }

            return result;
        }

        // Flight validation
        public Boolean IsActiveFlight(Guid segmentId)
        {
            Boolean result = false;
            foreach (FlightSegment mp in Segments)
            {
                if (mp.BookingSegmentId == segmentId)
                {
                    if (!string.IsNullOrEmpty(mp.FlightStatusRCD))
                    {
                        if (mp.FlightStatusRCD.Trim().ToUpper().Equals("ACTIVE"))
                        {
                            result = true;
                            break;
                        }
                    }
                }
            }
            return result;
        }
        // Flight validation get status if not active
        public string FlightStatus(Guid segmentId)
        {
            string _FlightStatus = string.Empty;
            
            var list = new List<KeyValuePair<string, string>>();
            list.Add(new KeyValuePair<string, string>( "F009","CANCELLED"));
            list.Add(new KeyValuePair<string, string>( "F008","DEPARTED"));
            list.Add(new KeyValuePair<string, string>( "F001","INACTIVE"));

            foreach (FlightSegment mp in Segments)
            {
                if (mp.BookingSegmentId == segmentId)
                {
                    _FlightStatus = mp.FlightStatusRCD;
                }
            }
            
            if (!string.IsNullOrEmpty(_FlightStatus))
            {
                // find list
                for (int i = 0; i < list.Count;i++ )
                {
                    if (_FlightStatus.ToUpper() == list[i].Value.ToUpper())
                    {
                        _FlightStatus = list[i].Key +","+ list[i].Value.ToLower();
                        break;
                    }
                }
            }
            return _FlightStatus;
        }

        // Flight validation
        public Boolean IsActiveFlight()
        {
            Boolean result = false;
            foreach (FlightSegment mp in Segments)
            {
                if (!string.IsNullOrEmpty(mp.FlightStatusRCD) && mp.FlightStatusRCD.Trim().ToUpper().Equals("ACTIVE"))
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        public Boolean IsNotValidFlightCheckIn(Guid segmentId)
        {
            Boolean result = false;

            foreach (FlightSegment fs in Segments)
            {
                if (fs.BookingSegmentId == segmentId && !string.IsNullOrEmpty(fs.FlightCheckInStatusRcd) && fs.FlightCheckInStatusRcd.ToUpper().Equals("FLOWN"))
                {
                    result = true;
                    break;
                }
                else if (fs.BookingSegmentId == segmentId && !string.IsNullOrEmpty(fs.FlightCheckInStatusRcd)  && fs.FlightCheckInStatusRcd.ToUpper().Equals("DISPATCHED"))
                {
                    result = true;
                    break;
                }
                else if (fs.BookingSegmentId == segmentId && !string.IsNullOrEmpty(fs.FlightCheckInStatusRcd) && fs.FlightCheckInStatusRcd.ToUpper().Equals("CLOSED"))
                {
                    result = true;
                    break;
                }
            }

            return result;
        }
        public Boolean IsFlownFlight(Guid segmentId)
        {
            Boolean result = false;

            foreach (FlightSegment fs in Segments)
            {
                if (fs.BookingSegmentId == segmentId && fs.FlightFlownDateTime != DateTime.MinValue)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        public Boolean IsFlownFlight()
        {
            Boolean result = false;

            foreach (FlightSegment fs in Segments)
            {
                if (fs.FlightFlownDateTime != DateTime.MinValue)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        public Boolean IsConnectingFlight()
        {
            Boolean result = false;
            foreach (Mapping mp in Mappings)
            {
                if (mp.FlightConnectionId != null && mp.FlightConnectionId != Guid.Empty)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        // Passenger Check in status by segment  for change flight only
        public Boolean IsCheckedPassenger(Guid segmentId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("CHECKED"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsBoardedPassenger(Guid segmentId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("BOARDED"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsFlownPassenger(Guid segmentId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("FLOWN"));

            if (total > 0)
                result = true;

            return result;
        }

        // Passenger Check in status
        public Boolean IsCheckedPassenger(Guid segmentId, Guid passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && item.PassengerId == passengerId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("CHECKED"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsBoardedPassenger(Guid segmentId, Guid passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && item.PassengerId == passengerId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("BOARDED"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsFlownPassenger(Guid segmentId, Guid passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && item.PassengerId == passengerId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("FLOWN"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsNoShowPassenger(Guid segmentId, Guid passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && item.PassengerId == passengerId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("NOSHOW"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsOffLoadedPassenger(Guid segmentId, Guid passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.BookingSegmentId == segmentId && item.PassengerId == passengerId && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("OFFLOADED"));

            if (total > 0)
                result = true;

            return result;
        }

        // Passenger Check in status by passenger now not used
        public Boolean IsCheckedPassenger(string passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.PassengerId == new Guid(passengerId) && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("CHECKED"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsBoardedPassenger(string passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.PassengerId == new Guid(passengerId) && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("BOARDED"));

            if (total > 0)
                result = true;

            return result;
        }
        public Boolean IsFlownPassenger(string passengerId)
        {
            Boolean result = false;
            Int32 total = Mappings.Count(item => item.PassengerId == new Guid(passengerId) && (item.PassengerCheckinStatusRcd != null) && item.PassengerCheckinStatusRcd.Trim().ToUpper().Equals("FLOWN"));

            if (total > 0)
                result = true;

            return result;
        }

        // Passenger status 
        public Boolean IsPassengerInBooking(String passengerId)
        {
            Boolean result = false;
            try
            {
                Int32 total = Mappings.Count(item => item.PassengerId == Guid.Parse(passengerId));

                if (total > 0)
                    result = true;
            }
            catch
            {
                throw;
            }

            return result;
        }
        public Boolean IsValidPassengerStatus(string passengerId, string segmentId)
        {
            Boolean result = false;
            foreach (Mapping mp in Mappings)
            {
                if (mp.PassengerId == new Guid(passengerId) && mp.BookingSegmentId == new Guid(segmentId))
                {
                    if (!string.IsNullOrEmpty(mp.PassengerStatusRcd) && mp.PassengerStatusRcd.ToUpper().Equals("OK"))
                    {
                        result = true;
                        break;
                    }
                }
            }
            return result;
        }
        public Boolean IsValidPassengerStatus(Guid passengerId)
        {
            Boolean result = false;
            foreach (Mapping mp in Mappings)
            {
                if (mp.PassengerId == passengerId )
                {
                    if (!string.IsNullOrEmpty(mp.PassengerStatusRcd) && mp.PassengerStatusRcd.ToUpper().Equals("OK"))
                    {
                        result = true;
                        break;
                    }
                }
            }
            return result;
        }

        // Passenger Infant 
        public Boolean FoundPassengerInfant(Guid passengerId)
        {
            Boolean result = false;
            foreach (Mapping mp in Mappings)
            {
                if (mp.PassengerId == passengerId)
                {
                    if (!string.IsNullOrEmpty(mp.PassengerTypeRcd) && mp.PassengerTypeRcd.ToUpper().Equals("INF"))
                    {
                        result = true;
                        break;
                    }
                }
            }
            return result;
        }

        // Duplicate Passenger 
        public Boolean DuplicatePassengerName(string name, Guid passengerId)
        {
            Boolean result = false;
            foreach (Passenger p in Passengers)
            {
                if ((p.TitleRcd +p.Firstname + p.Middlename + p.Lastname).Equals(name.ToUpper()) && p.PassengerId != passengerId)
                {
                        result = true;
                        break;
                }
            }
            return result;
        }

        // SSR
        public Boolean IsDuplicateSSR(String serviceCode, Guid passengerId, Guid segmentId)
        {
            Boolean result = false;
            if (Services != null && Services.Count > 0)
            {
                foreach (Avantik.Web.Service.Entity.Booking.PassengerService entitySSR in Services)
                {
                    if (!string.IsNullOrEmpty(entitySSR.SpecialServiceRcd) &&  !string.IsNullOrEmpty(serviceCode))
                    {
                        if (entitySSR.SpecialServiceRcd.Trim().ToUpper() == serviceCode.Trim().ToUpper() && entitySSR.PassengerId == passengerId && entitySSR.BookingSegmentId == segmentId)
                        {
                            result = true;
                        }
                    }
                }
            }

            // check fee as SSR fee
            if (result == false)
            {
                if (Fees != null && Fees.Count > 0)
                {
                    foreach (Avantik.Web.Service.Entity.Booking.Fee f in Fees)
                    {
                        if (!string.IsNullOrEmpty(f.FeeRcd) && !string.IsNullOrEmpty(serviceCode))
                        {
                            if (f.FeeRcd.Trim().ToUpper() == serviceCode.Trim().ToUpper() && f.BookingSegmentId == segmentId && f.PassengerId == passengerId)
                            {
                                result = true;
                            }
                        }
                    }
                }
            }

            return result;
        }

        //Ticket
        public Boolean IsFoundTicketNumber(string segmentId, string passengerId)
        {
            Boolean result = false;
            if (Mappings != null && Mappings.Count > 0)
            {
                foreach (Avantik.Web.Service.Entity.Booking.Mapping m in Mappings)
                {
                    if (m.BookingSegmentId == new Guid(segmentId) && m.PassengerId == new Guid(passengerId))
                    {
                       if(!string.IsNullOrEmpty(m.TicketNumber))
                       {
                           result = true;
                       }
                    }
                }
            }

            return result;
        }

        #endregion

    }
}
