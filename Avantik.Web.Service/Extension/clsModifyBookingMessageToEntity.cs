using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.ManageBooking;

namespace Avantik.Web.Service.Extension
{
    public static class ModifyBookingMessageToEntity
    {
        public static Flight ToEntity(this Message.ManageBooking.ChangeFlight modifyFlight)
        {
            Flight flight = null;
            if (modifyFlight != null)
            {
                flight = new Flight();

                flight.BoardingClassRcd = modifyFlight.NewSegment.BoardingClassRcd;
                flight.BookingClassRcd = modifyFlight.NewSegment.BookingClassRcd;
                flight.FlightId = new Guid(modifyFlight.NewSegment.FlightId);
                flight.OriginRcd = modifyFlight.NewSegment.OriginRcd;
                flight.DepartureDate = modifyFlight.NewSegment.DepartureDate;
                flight.FairId = new Guid(modifyFlight.NewSegment.FareId);
                flight.EticketFlag = 1;

                flight.OdOriginRcd = modifyFlight.NewSegment.OdOriginRcd;
                flight.OdDestinationRcd = modifyFlight.NewSegment.OdDestinationRcd;

                if (!string.IsNullOrEmpty(modifyFlight.NewSegment.TransitFlightId))
                {
                    if (!modifyFlight.NewSegment.TransitFlightId.ToString().Equals(Guid.Empty.ToString()))
                    {
                        flight.FlightConnectionId = new Guid(modifyFlight.NewSegment.TransitFlightId);
                        flight.DestinationRcd = modifyFlight.NewSegment.TransitAirportRcd;
                    }
                    else
                    {
                        flight.DestinationRcd = modifyFlight.NewSegment.DestinationRcd;
                    }
                }

                flight.ExchangedSegmentId = Guid.Empty;

            }
            return flight;
        }

        public static Flight TransitToEntity(this Message.ManageBooking.ChangeFlight modifyFlight)
        {
            Flight flight = null;
            if (modifyFlight != null)
            {
                flight = new Flight();
                flight.BoardingClassRcd = modifyFlight.NewSegment.TransitBoardingClassRcd;
                flight.BookingClassRcd = modifyFlight.NewSegment.TransitBookingClassRcd;
                flight.DestinationRcd = modifyFlight.NewSegment.DestinationRcd;
                flight.FlightId = new Guid(modifyFlight.NewSegment.TransitFlightId);
                flight.OriginRcd = modifyFlight.NewSegment.TransitAirportRcd;
                flight.DepartureDate = modifyFlight.NewSegment.TransitDepartureDate;
                flight.FairId = new Guid(modifyFlight.NewSegment.TransitFareId);
                flight.OdOriginRcd = modifyFlight.NewSegment.OriginRcd;
                flight.OdDestinationRcd = modifyFlight.NewSegment.DestinationRcd;
                flight.FlightConnectionId = new Guid(modifyFlight.NewSegment.TransitFlightId);
                flight.EticketFlag = 1;
            }
            return flight;
        }

        public static IList<Flight> ToListEntity(this  IList<Message.ManageBooking.ChangeFlight> modifyFlights)
        {
            IList<Flight> segmentList = null;

            if (modifyFlights != null)
            {
                segmentList = new List<Flight>();

                for (int i = 0; i < modifyFlights.Count; i++)
                {
                    segmentList.Add(modifyFlights[i].ToEntity());

                    // for connection flight
                    if (!string.IsNullOrEmpty(modifyFlights[i].NewSegment.TransitFlightId))
                    {
                        if (!modifyFlights[i].NewSegment.TransitFlightId.ToString().Equals(Guid.Empty.ToString()))
                        {
                            segmentList.Add(modifyFlights[i].TransitToEntity());
                        }
                    }

                }
            }

            return segmentList;
        }

        public static IList<Entity.Booking.Payment> ToListEntity(this  IList<Message.ManageBooking.Payment> pay, string userId, string bookingId)
        {
            IList<Entity.Booking.Payment> paymentList = null;

            if (pay != null)
            {
                paymentList = new List<Entity.Booking.Payment>();
                for (int i = 0; i < pay.Count; i++)
                {
                    paymentList.Add(pay[i].ToEntity(userId, bookingId));
                }
            }
            return paymentList;
        }
        public static Entity.Booking.Payment ToEntity(this  Message.ManageBooking.Payment pay, string userId, string bookingId)
        {
            Entity.Booking.Payment payment = null;
            if (pay != null)
            {
                payment = new Entity.Booking.Payment();
                payment.BookingPaymentId = Guid.NewGuid();

                payment.BookingId = Guid.Parse(bookingId);
               // payment.VoucherPaymentId = Guid.NewGuid();
                payment.FormOfPaymentRcd = pay.FormOfPaymentRcd.Trim().ToUpper();

                // when pay with voucher will gen id at payment function

                //if (payment.FormOfPaymentRcd == "VOUCHER")
                //{
                //   payment.VoucherPaymentId = Guid.NewGuid();
                //}

               // payment.CurrencyRcd = pay.CurrencyRcd;
              //  payment.ReceiveCurrencyRcd = pay.ReceiveCurrencyRcd;
                //payment.AgencyCode = pay.AgencyCode;
                //payment.DebitAgencyCode = pay.DebitAgencyCode;
             //    payment.PaymentAmount = pay.PaymentAmount;
                //payment.ReceivePaymentAmount = pay.ReceivePaymentAmount;
                //payment.AcctPaymentAmount = pay.AcctPaymentAmount;
                payment.PaymentBy = Guid.Parse(userId);
                payment.PaymentDateTime = DateTime.Now;
              //  payment.PaymentDueDateTime = pay.PaymentDueDateTime;
                //payment.DocumentAmount = pay.DocumentAmount;
                payment.ExpiryMonth = pay.ExpiryMonth;
                payment.ExpiryYear = pay.ExpiryYear;
              //  payment.RecordLocator = pay.RecordLocator;
                payment.CvvCode = pay.CvvCode;
                payment.NameOnCard = pay.NameOnCard;
                payment.DocumentNumber = pay.DocumentNumber;
                payment.DocumentPassword = pay.DocumentPassword;
                payment.FormOfPaymentSubtypeRcd = pay.FormOfPaymentSubtypeRcd.Trim().ToUpper();
                payment.City = pay.City;
                payment.State = pay.State;
                payment.Street = pay.Street;
              //  payment.PoBox = pay.PoBox;
                payment.AddressLine1 = pay.AddressLine1;
              //  payment.AddressLine2 = pay.AddressLine2;
              //  payment.District = pay.District;
               // payment.Province = pay.Province;
                payment.ZipCode = pay.ZipCode;
                payment.CountryRcd = pay.CountryRcd;
                payment.CreateBy = Guid.Parse(userId);
                payment.CreateDateTime = DateTime.Now;
                payment.UpdateBy = Guid.Parse(userId);
                payment.UpdateDateTime = DateTime.Now;
               // payment.PosIndicator = pay.PosIndicator;
                payment.IssueMonth = pay.IssueMonth;
                payment.IssueYear = pay.IssueYear;
                payment.IssueNumber = pay.IssueNumber;
              //  payment.TransactionReference = pay.TransactionReference;
               // payment.ApprovalCode = pay.ApprovalCode;
               // payment.BankName = pay.BankName;
               // payment.BankCode = pay.BankCode;
              //  payment.BankIban = pay.BankIban;
                if (!string.IsNullOrEmpty(pay.IpAddress))
                {
                    payment.IpAddress = pay.IpAddress;
                }
                else
                {
                    payment.IpAddress = "127.0.0.1";
                }
              //  payment.PaymentReference = pay.PaymentReference;
                //payment.AllocatedAmount = pay.AllocatedAmount;
                //payment.PaymentTypeRcd = pay.PaymentTypeRcd;
              //  payment.RefundedAmount = pay.RefundedAmount;
               // payment.PaymentNumber = pay.PaymentNumber;
             //   payment.PaymentRemark = pay.PaymentRemark;
            }

            return payment;
        }

        public static IList<Entity.SeatAssign> ToListEntity(this  IList<Message.ManageBooking.SeatAssign> s)
        {
            IList<Entity.SeatAssign> seatList = null;

            if (s != null)
            {
                seatList = new List<Entity.SeatAssign>();
                for (int i = 0; i < s.Count; i++)
                {
                    seatList.Add(s[i].ToEntity());
                }
            }
            return seatList;
        }
        public static Entity.SeatAssign ToEntity(this  Message.ManageBooking.SeatAssign s)
        {
            Entity.SeatAssign seat = null;
            if (s != null)
            {
                seat = new Entity.SeatAssign();
                seat.BookingSegmentID = s.BookingSegmentID;
                seat.PassengerID = s.PassengerID;
                seat.SeatColumn = s.SeatColumn.Trim().ToUpper();
                seat.SeatRow = s.SeatRow;
                seat.SeatNumber = s.SeatRow +  s.SeatColumn.Trim().ToUpper();

                if (!string.IsNullOrEmpty(s.SeatFeeRcd))
                {
                    seat.SeatFeeRcd = s.SeatFeeRcd.Trim().ToUpper();
                }
            }

            return seat;
        }

        //public static IList<Message.ManageBooking.SeatAssign> ToListEntity2(this  IList<Message.ManageBooking.SeatAssign> s)
        //{
        //    IList<Message.ManageBooking.SeatAssign> seatList = null;

        //    if (s != null)
        //    {
        //        seatList = new List<Message.ManageBooking.SeatAssign>();
        //        for (int i = 0; i < s.Count; i++)
        //        {
        //            seatList.Add(s[i].ToEntity2());
        //        }
        //    }
        //    return seatList;
        //}
        //public static Message.ManageBooking.SeatAssign ToEntity2(this  Message.ManageBooking.SeatAssign s)
        //{
        //    Message.ManageBooking.SeatAssign seat = null;
        //    if (s != null)
        //    {
        //        seat = new Message.ManageBooking.SeatAssign();
        //        seat.BookingSegmentID = s.BookingSegmentID;
        //        seat.PassengerID = s.PassengerID;
        //        seat.SeatColumn = s.SeatColumn.Trim().ToUpper();
        //        seat.SeatRow = s.SeatRow;

        //        seat.SeatNumber = s.SeatRow + s.SeatColumn.Trim().ToUpper();

        //        if (!string.IsNullOrEmpty(s.SeatFeeRcd))
        //        {
        //            seat.SeatFeeRcd = s.SeatFeeRcd.Trim().ToUpper();
        //        }
        //    }

        //    return seat;
        //}


        public static IList<Entity.NameChange> ToListEntity(this  IList<Message.ManageBooking.NameChange> n)
        {
            IList<Entity.NameChange> nameList = null;

            if (n != null)
            {
                nameList = new List<Entity.NameChange>();
                for (int i = 0; i < n.Count; i++)
                {
                    nameList.Add(n[i].ToEntity());
                }
            }
            return nameList;
        }
        public static Entity.NameChange ToEntity(this  Message.ManageBooking.NameChange n)
        {
            Entity.NameChange entityName = null;
            if (n != null)
            {
                entityName = new Entity.NameChange();
                entityName.PassengerId = new Guid(n.PassengerId);
                entityName.TitleRcd = n.TitleRcd;
                entityName.Firstname = n.Firstname;
                entityName.Middlename = n.Middlename;
                entityName.Lastname = n.Lastname;
                entityName.DateOfBirth = n.DateOfBirth;
                entityName.GenderTypeRcd = n.GenderTypeRcd;
            }

            return entityName;
        }

    }
}
