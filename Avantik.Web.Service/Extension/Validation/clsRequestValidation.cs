using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.ManageBooking;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Message.SeatMap;
using System.Net.Mail;
using System.Collections;
using System.Text.RegularExpressions;

namespace Avantik.Web.Service.Extension
{
    public static class RequestValidation
    {
        public static InitializeResponse Valid(this Avantik.Web.Service.Message.InitializeRequest request, InitializeResponse response)
        {
            bool bError = false;
            string strMessage = "ServiceInitialize Error: ";
            string code = string.Empty;

            if (string.IsNullOrEmpty(request.AgencyCode))
            {
                bError = true;
                strMessage = "AgencyCode is required.";
                code = "A011";
            }
            else if (string.IsNullOrEmpty(request.AgencyCode.Trim()))
            {
                bError = true;
                strMessage = "AgencyCode is required.";
                code = "A011";
            }
            else if (string.IsNullOrEmpty(request.AgentLogon))
            {
                bError = true;
                strMessage = "AgentLogon is required.";
                code = "A012";
            }
            else if (string.IsNullOrEmpty(request.AgentLogon.Trim()))
            {
                bError = true;
                strMessage = "AgentLogon is required.";
                code = "A012";
            }
            else if (string.IsNullOrEmpty(request.AgentPassword))
            {
                bError = true;
                strMessage = "AgentPassword is required.";
                code = "A013";
            }
            else if (string.IsNullOrEmpty(request.AgentPassword.Trim()))
            {
                bError = true;
                strMessage = "AgentPassword is required.";
                code = "A013";
            }
            if (bError)
            {
                response.Code = code;
                response.Message = strMessage;
                response.Success = false;
            }
            else
            {
                response.Code = "000";
                response.Message = "Success";
                response.Success = true;
            }

            return response;
        }

        public static CalculateSeatFeesResponse Valid(this IList<Avantik.Web.Service.Message.ManageBooking.SeatAssign> seats, CalculateSeatFeesResponse response, Booking booking,Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Seat Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowSeat == "1")
            {
                for (int i = 0; i < seats.Count; i++)
                {
                    if (string.IsNullOrEmpty(seats[i].BookingSegmentID))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is required.";
                        code = "F007";
                    }
                    else if (string.IsNullOrEmpty(seats[i].BookingSegmentID.Trim()))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is required.";
                        code = "F007";
                    }
                    else if (string.IsNullOrEmpty(seats[i].PassengerID))
                    {
                        bError = true;
                        strMessage += "PassengerID is required.";
                        code = "L007";
                    }
                    else if (string.IsNullOrEmpty(seats[i].PassengerID.Trim()))
                    {
                        bError = true;
                        strMessage += "PassengerID is required.";
                        code = "L007";
                    }

                    else if (booking.IsPassengerInBooking(seats[i].PassengerID.ToString()) == false)
                    {
                        bError = true;
                        strMessage += "Passenger is not in booking.";
                        code = "L006";
                    }
                    else if (booking.IsSegmentInBooking(seats[i].BookingSegmentID.ToString()) == false)
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is not in booking.";
                        code = "F006";
                    }
                    // validate segment
                    else if (!booking.IsValidSegmentStatus(seats[i].BookingSegmentID.ToString()))
                    {
                        bError = true;
                        strMessage += "Segment status is invalid.";
                        code = "F005";
                    }

                    //IsValidFlighttatus
                    else if (!booking.IsActiveFlight(new Guid(seats[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Flight is inactive.";
                        code = "F001";
                    }
                    else if (booking.IsNotValidFlightCheckIn(new Guid(seats[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "FlightCheckInStatus is invalid.";
                        code = "F002";
                    }
                    else if (booking.IsFlownFlight(new Guid(seats[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Flight is flown.";
                        code = "F003";
                    }
                   // Passenger Check in status 
                    else if (booking.IsCheckedPassenger(new Guid(seats[i].BookingSegmentID), new Guid(seats[i].PassengerID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsBoardedPassenger(new Guid(seats[i].BookingSegmentID), new Guid(seats[i].PassengerID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsFlownPassenger(new Guid(seats[i].BookingSegmentID), new Guid(seats[i].PassengerID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsNoShowPassenger(new Guid(seats[i].BookingSegmentID), new Guid(seats[i].PassengerID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsOffLoadedPassenger(new Guid(seats[i].BookingSegmentID), new Guid(seats[i].PassengerID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }

                    else if (string.IsNullOrEmpty(seats[i].SeatColumn))
                    {
                        bError = true;
                        strMessage += "SeatColumn is required.";
                        code = "P021";
                    }
                    else if (string.IsNullOrEmpty(seats[i].SeatColumn.Trim()))
                    {
                        bError = true;
                        strMessage += "SeatColumn is required.";
                        code = "P021";
                    }
                    else if (seats[i].SeatRow.ToString().Length == 0)
                    {
                        bError = true;
                        strMessage += "SeatRow is required.";
                        code = "P020";
                    }
                    else if (booking.IsFlightStatusAllowAssignSeat(seats[i].BookingSegmentID) == false)
                    {
                        bError = true;
                        strMessage += "Flight is not allowed seat.";
                        code = "P027";
                    }
                     // dup in flight

                    // dup seat in booking
                    //else if (booking.IsDuplicateSeat(seats[i].SeatRow + seats[i].SeatColumn) == true)
                    //{
                    //    bError = true;
                    //    strMessage += "Duplicated Seat.";
                    //    code = "P026";
                    //}
                    //else if (seats.Count > 1 && IsDuplicateSeat(seats, seats[i].SeatRow + seats[i].SeatColumn))
                    //{
                    //    bError = true;
                    //    strMessage += "Duplicated Seat.";
                    //    code = "P026";
                    //}

                    if (bError)
                    {
                        response.Code = code;
                        response.Message = strMessage;
                        response.Success = false;
                        break;
                    }
                    else
                    {
                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }
                }
            }
            else
            {
                response.Code = "A008";
                response.Message = "Agency is not allowed seat.";
                response.Success = false;
            }
            
            return response;

        }

        public static NameChangeResponse Valid(this IList<Avantik.Web.Service.Message.ManageBooking.NameChange> name, NameChangeResponse response, Entity.Authentication objAuthen, Booking booking)
        {
            bool bError = false;
            string strMessage = "NameChange Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowNameChange == "1")
            {
                for (int i = 0; i < name.Count; i++)
                {
                    if (string.IsNullOrEmpty(name[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "PassengerId is required.";
                        code = "L007";
                    }
                    else if (!Helpers.DataType.IsGuid(name[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "PassengerId is invalid.";
                        code = "L011";
                    }
                    else if (!booking.IsPassengerInBooking(name[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "Passenger is not in booking.";
                        code = "L006";
                    }
                    // check only status OK
                    else if (!booking.IsValidPassengerStatus(new Guid(name[i].PassengerId)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    // validate segment  ps =OK  seg =HK
                    else if (!booking.IsValidSegmentStatus())
                    {
                        bError = true;
                        strMessage += "Segment status is invalid.";
                        code = "F005";
                    }

                    //IsValidFlighttatus active
                    else if (!booking.IsActiveFlight())
                    {
                        bError = true;
                        strMessage += "Flight status is invalid.";
                        code = "F010";
                    }
                        //flight flown
                    else if (booking.IsFlownFlight())
                    {
                        bError = true;
                        strMessage += "Flight is flown.";
                        code = "F003";
                    }
                    // Passenger Check in status 
                    else if (booking.IsCheckedPassenger(name[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsBoardedPassenger(name[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsFlownPassenger(name[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (!ValidGender(name[i].GenderTypeRcd))
                    {
                        bError = true;
                        strMessage += "GenderType is invalid.";
                        code = "L010";
                    }
                    else if (!Regex.IsMatch(name[i].Firstname + name[i].Middlename + name[i].Lastname, "^[a-zA-Z]*$"))
                    {
                        bError = true;
                        strMessage += "Passenger name is invalid.";
                        code = "L011";
                    }
                    else if (booking.DuplicatePassengerName(name[i].TitleRcd +name[i].Firstname + name[i].Middlename + name[i].Lastname, new Guid(name[i].PassengerId)))
                    {
                        bError = true;
                        strMessage += "Duplicated Passenger name.";
                        code = "L013";
                    }

                    else if (IsDuplicateName(name))
                    {
                        bError = true;
                        strMessage += "Duplicated Passenger name.";
                        code = "L013";
                    }

                    if (bError)
                    {
                        response.Code = code;
                        response.Message = strMessage;
                        response.Success = false;
                        break;
                    }
                    else
                    {
                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }
                }
            }
            else
            {
                response.Code = "A008";
                response.Message = "Agency is not allowed to change name.";
                response.Success = false;
            }

            return response;

        }

        private static bool IsDuplicateName(IList<Avantik.Web.Service.Message.ManageBooking.NameChange> name)
        {
            bool r = false;
            ArrayList al = new ArrayList();

            for (int i = 0; i < name.Count; i++)
            {
               // string midname = string.IsNullOrEmpty(name[i].Middlename) ? string.Empty : name[i].Middlename.ToUpper();
                string pn = (name[i].TitleRcd + name[i].Firstname + name[i].Middlename + name[i].Lastname).ToUpper();
                al.Add(pn);
            }
            string[] array = al.ToArray(typeof(string)) as string[];

            if (array.Distinct().Count() != array.Count())
            {
                r = true;
            }
            return r;
        }

        private static bool ValidGender(string GenderTypeRcd)
        {
            bool result = false;

                if (GenderTypeRcd != null)
                {
                    if (GenderTypeRcd.Trim().ToUpper() == "M" || GenderTypeRcd.Trim().ToUpper() == "F")
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }
                else
                {
                    result = false;
                }
            return result;
        }
    

    
        public static PassengerInfoResponse Valid(this IList<Avantik.Web.Service.Message.PassengerInfo> pf, PassengerInfoResponse response, Booking booking,Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Updated Passenger Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowInFoPassenger == "1")
            {
                for (int i = 0; i < pf.Count; i++)
                {
                    if (booking.IsLockBooking())
                    {
                        bError = true;
                        strMessage += "Booking is Locked.";
                        code = "L001";
                    }
                    else if (string.IsNullOrEmpty(pf[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "PassengerId required.";
                        code = "L007";
                    }
                    else if (!Helpers.DataType.IsGuid(pf[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "PassengerId is invalid.";
                        code = "L011";
                    }
                    else if (booking.IsPassengerInBooking(pf[i].PassengerId) == false)
                    {
                        bError = true;
                        strMessage += "Passenger is not in booking.";
                        code = "L006";
                    }
                    // check only status OK
                    else if (!booking.IsValidPassengerStatus(new Guid(pf[i].PassengerId)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }

                    // length
                    else if (!string.IsNullOrEmpty(pf[i].PassportBirthPlace) && pf[i].PassportBirthPlace.Length > 60)
                    {
                        bError = true;
                        strMessage += "PassportBirthPlace is over length.";
                        code = "I002";
                    }
                    else if (!string.IsNullOrEmpty(pf[i].PassportIssuePlace) && pf[i].PassportIssuePlace.Length > 60)
                    {
                        bError = true;
                        strMessage += "PassportIssuePlace is over length.";
                        code = "I003";
                    }

                    else if (!string.IsNullOrEmpty(pf[i].DocumentNumber) && pf[i].DocumentNumber.Length > 60)
                    {
                        bError = true;
                        strMessage += "DocumentNumber is over length.";
                        code = "I004";
                    }
                    else if (!string.IsNullOrEmpty(pf[i].DocumentTypeRcd) && pf[i].DocumentTypeRcd.Length > 10)
                    {
                        bError = true;
                        strMessage += "DocumentTypeRcd is over length.";
                        code = "I005";
                    }
                    else if (!string.IsNullOrEmpty(pf[i].NationalityRcd) && pf[i].NationalityRcd.Length > 10)
                    {
                        bError = true;
                        strMessage += "NationalityRcd is over length.";
                        code = "I006";
                    }
                    else if (!string.IsNullOrEmpty(pf[i].PassportIssueCountryRcd) && pf[i].PassportIssueCountryRcd.Length > 10)
                    {
                        bError = true;
                        strMessage += "PassportIssueCountryRcd is over length.";
                        code = "I007";
                    }

                    if (bError)
                    {
                        response.Code = code;
                        response.Message = strMessage;
                        response.Success = false;
                        break;
                    }
                    else
                    {
                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }
                }
            }
            else
            {
                response.Code = "A008";
                response.Message = "Agent is not allowed change passenger infomation.";
                response.Success = false;
            }

            return response;

        }

        // for modify sent numberOfUnit
        public static BaggageFeesResponse Valid(this IList<Avantik.Web.Service.Message.ManageBooking.BaggageRequest> bags, BaggageFeesResponse response, Booking booking,Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Baggage Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowService == "1")
            {
                for (int i = 0; i < bags.Count; i++)
                {
                    if (string.IsNullOrEmpty(bags[i].BookingSegmentID))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is required.";
                        code = "F007";
                    }
                    else if (string.IsNullOrEmpty(bags[i].BookingSegmentID.Trim()))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is required.";
                        code = "F007";
                    }
                    else if (string.IsNullOrEmpty(bags[i].PassengerID))
                    {
                        bError = true;
                        strMessage += "PassengerID is required.";
                        code = "L007";
                    }
                    else if (string.IsNullOrEmpty(bags[i].PassengerID.Trim()))
                    {
                        bError = true;
                        strMessage += "PassengerID is required.";
                        code = "L007";
                    }
                    else if (!booking.IsPassengerInBooking(bags[i].PassengerID.Trim()))
                    {
                        bError = true;
                        strMessage += "Passenger is not in Booking.";
                        code = "L006";
                    }
                    else if (!booking.IsSegmentInBooking(bags[i].BookingSegmentID.Trim()))
                    {
                        bError = true;
                        strMessage += "Booking Segment is not in Booking.";
                        code = "F006";
                    }

                    // validate segment
                    else if (!booking.IsValidSegmentStatus(bags[i].BookingSegmentID.ToString()))
                    {
                        bError = true;
                        strMessage += "Segment status is invalid.";
                        code = "F005";
                    }

                    //IsValidFlighttatus
                    else if (!booking.IsActiveFlight(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Flight is InActive.";
                        code = "F001";
                    }
                    else if (booking.IsNotValidFlightCheckIn(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "FlightCheckInStatus is invalid.";
                        code = "F002";
                    }
                    else if (booking.IsFlownFlight(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Flight is flown.";
                        code = "F003";
                    }

                   // Passenger Check in status 
                    else if (booking.IsCheckedPassenger(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsBoardedPassenger(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsFlownPassenger(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }

                    else if (bags[i].NumberOfUnit.ToString().Length == 0)
                    {
                        bError = true;
                        strMessage += "NumberOfUnit is required.";
                        code = "V014";
                    }
                    // IsDecimal
                    else if (!IsDecimal(bags[i].NumberOfUnit.ToString()))
                    {
                        bError = true;
                        strMessage += "NumberOfUnit should be numberic.";
                        code = "V015";
                    }
                    else if (bags[i].NumberOfUnit == 0)
                    {
                        bError = true;
                        strMessage += "NumberOfUnit should be more than 0.";
                        code = "V016";
                    }
                    // FoundPassengerInfat
                    else if (booking.FoundPassengerInfant(new Guid(bags[i].PassengerID)))
                    {
                        bError = true;
                        strMessage += "Infant is not allowed to book Baggage.";
                        code = "L008";
                    }

                    if (bError)
                    {
                        response.Code = code;
                        response.Message = strMessage;
                        response.Success = false;
                        break;
                    }
                    else
                    {
                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }

                }
            }
            else
            {
                response.Code = "A008";
                response.Message = "Agency is not allowed Baggage.";
                response.Success = false;
            }

            return response;

        }

        // for view fee
        public static BaggageFeesResponse Valid(this IList<Avantik.Web.Service.Message.ManageBooking.Baggage> bags, BaggageFeesResponse response, Booking booking, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Baggage Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowService == "1")
            {
                for (int i = 0; i < bags.Count; i++)
                {
                    if (string.IsNullOrEmpty(bags[i].BookingSegmentID))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is required.";
                        code = "F007";
                    }
                    if (string.IsNullOrEmpty(bags[i].BookingSegmentID.Trim()))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is required.";
                        code = "F007";
                    }

                    else if (string.IsNullOrEmpty(bags[i].PassengerID))
                    {
                        bError = true;
                        strMessage += "PassengerID is required.";
                        code = "L007";
                    }
                    else if (string.IsNullOrEmpty(bags[i].PassengerID.Trim()))
                    {
                        bError = true;
                        strMessage += "PassengerID is required.";
                        code = "L007";
                    }

                    else if (!booking.IsPassengerInBooking(bags[i].PassengerID.Trim()))
                    {
                        bError = true;
                        strMessage += "Passenger is not in Booking.";
                        code = "L006";
                    }
                    else if (!booking.IsSegmentInBooking(bags[i].BookingSegmentID.Trim()))
                    {
                        bError = true;
                        strMessage += "Booking Segment is not in Booking.";
                        code = "F006";
                    }
                    // validate segment
                    else if (!booking.IsValidSegmentStatus(bags[i].BookingSegmentID.ToString()))
                    {
                        bError = true;
                        strMessage += "Segment status invalid.";
                        code = "F005";
                    }

                    //IsValidFlighttatus
                    else if (!booking.IsActiveFlight(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Flight is InActive.";
                        code = "F001";
                    }
                    else if (booking.IsNotValidFlightCheckIn(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "FlightCheckInStatus is invalid.";
                        code = "F002";
                    }
                    else if (booking.IsFlownFlight(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Flight is flown.";
                        code = "F003";
                    }

                   // Passenger Check in status 
                    else if (booking.IsCheckedPassenger(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsBoardedPassenger(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsFlownPassenger(new Guid(bags[i].BookingSegmentID)))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }

                    if (bError)
                    {
                        response.Code = code;
                        response.Message = strMessage;
                        response.Success = false;
                        break;
                    }
                    else
                    {
                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }
                }
            }
            else
            {
                response.Code = "A008";
                response.Message = "Agency is not allowed Baggage.";
                response.Success = false;
            }

            return response;

        }

        // use for modify
        public static CalculateSpecialServiceFeesResponse Valid(this IList<Avantik.Web.Service.Message.SSR.Service> services, CalculateSpecialServiceFeesResponse response, Booking booking,Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Service Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowService == "1")
            {
                for (int i = 0; i < services.Count; i++)
                {
                    if (string.IsNullOrEmpty(services[i].BookingSegmentId))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is required.";
                        code = "F007";
                    }

                    else if (!Helpers.DataType.IsGuid(services[i].BookingSegmentId))
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is invalid.";
                        code = "L012";
                    }
                    else if (string.IsNullOrEmpty(services[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "PassengerID is required.";
                        code = "L007";
                    }
                    else if (!Helpers.DataType.IsGuid(services[i].PassengerId))
                    {
                        bError = true;
                        strMessage += "PassengerId is invalid.";
                        code = "L011";
                    }
                    else if (string.IsNullOrEmpty(services[i].SpecialServiceCode))
                    {
                        bError = true;
                        strMessage += "SpecialServiceRcd is required.";
                        code = "V009";
                    }

                    else if (string.IsNullOrEmpty(services[i].SpecialServiceCode.Trim()))
                    {
                        bError = true;
                        strMessage += "SpecialServiceRcd is required.";
                        code = "V009";
                    }
                    else if (booking.IsPassengerInBooking(services[i].PassengerId.ToString()) == false)
                    {
                        bError = true;
                        strMessage += "Passenger is not in booking.";
                        code = "L006";
                    }
                    else if (booking.IsSegmentInBooking(services[i].BookingSegmentId.ToString()) == false)
                    {
                        bError = true;
                        strMessage += "BookingSegmentID is not in booking.";
                        code = "F006";
                    }

                    // validate segment
                    else if (!booking.IsValidSegmentStatus(services[i].BookingSegmentId.ToString()))
                    {
                        bError = true;
                        strMessage += "Segment status is invalid.";
                        code = "F005";
                    }

                    //IsValidFlighttatus
                    else if (!booking.IsActiveFlight(new Guid(services[i].BookingSegmentId)))
                    {
                        bError = true;
                        strMessage += "Flight is InActive.";
                        code = "F001";
                    }
                    else if (booking.IsNotValidFlightCheckIn(new Guid(services[i].BookingSegmentId)))
                    {
                        bError = true;
                        strMessage += "FlightCheckInStatus is invalid.";
                        code = "F002";
                    }
                    else if (booking.IsFlownFlight(new Guid(services[i].BookingSegmentId)))
                    {
                        bError = true;
                        strMessage += "Flight is flown.";
                        code = "F003";
                    }

                   // Passenger Check in status 
                    else if (booking.IsCheckedPassenger(services[i].BookingSegmentId))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsBoardedPassenger(services[i].BookingSegmentId))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    else if (booking.IsFlownPassenger(services[i].BookingSegmentId))
                    {
                        bError = true;
                        strMessage += "Passenger status is invalid.";
                        code = "L009";
                    }
                    // FoundPassengerInfat
                    else if (booking.FoundPassengerInfant(new Guid(services[i].PassengerId)))
                    {
                        bError = true;
                        strMessage += "Infant is not allowed to book SSR.";
                        code = "L008";
                    }
                    // valid number
                    else if (services[i].NumberOfUnits.ToString().Length == 0)
                    {
                        bError = true;
                        strMessage += "NumberOfUnit is required.";
                        code = "V014";
                    }
                    else if (!IsDecimal(services[i].NumberOfUnits.ToString()))
                    {
                        bError = true;
                        strMessage += "NumberOfUnit should be numberic.";
                        code = "V015";
                    }
                    else if (services[i].NumberOfUnits == 0)
                    {
                        bError = true;
                        strMessage += "NumberOfUnit should be more than 0.";
                        code = "V016";
                    }
                    //else if (booking != null && booking.IsDuplicateSSR(services[i].SpecialServiceCode,new Guid(services[i].PassengerId),services[i].BookingSegmentId))
                    //{
                    //    bError = true;
                    //    strMessage += "Duplicated SSR.";
                    //    code = "V013";
                    //}
                    //
                    if (bError)
                    {
                        response.Code = code;
                        response.Message = strMessage;
                        response.Success = false;
                        break;
                    }
                    else
                    {
                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }
                }
            }
            else
            {
                response.Code = "A008";
                response.Message = "Agency is not allowed Service.";
                response.Success = false;
            }

            return response;

        }
                
        public static ModifyBookingResponse Valid(this IList<Avantik.Web.Service.Message.ManageBooking.Payment> pay, ModifyBookingResponse response, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Payment Error: ";
            for (int i = 0; i < pay.Count; i++)
            {
                //if (!string.IsNullOrEmpty(pay[i].PaymentNumber) && !IsAllDigits(pay[i].PaymentNumber))
                //{
                //    bError = true;
                //    strMessage += "PaymentNumber should be numberic.";
                //}
                if (string.IsNullOrEmpty(pay[i].FormOfPaymentRcd))
                {
                    bError = true;
                    strMessage += "FormOfPaymentRcd is required.";
                }
                else if (string.IsNullOrEmpty(pay[i].FormOfPaymentRcd.Trim()))
                {
                    bError = true;
                    strMessage += "FormOfPaymentRcd is required.";
                }
                else if (string.IsNullOrEmpty(pay[i].FormOfPaymentSubtypeRcd))
                {
                    bError = true;
                    strMessage += "FormOfPaymentSubtypeRcd is required.";
                }

                else if (string.IsNullOrEmpty(pay[i].FormOfPaymentSubtypeRcd.Trim()))
                {
                    bError = true;
                    strMessage += "FormOfPaymentSubtypeRcd is required.";
                }
                else if (pay[i].DocumentNumber.Length == 0)
                {
                    bError = true;
                    strMessage += "DocumentNumber is required.";
                }
                //else if (pay[i].PaymentAmount.ToString().Length == 0)
                //{
                //    bError = true;
                //    strMessage += "PaymentAmount is required.";
                //}
                //else if (!IsDecimal(pay[i].PaymentAmount.ToString()))
                //{
                //    bError = true;
                //    strMessage += "PaymentAmount should be numberic.";
                //}

                if (bError)
                {
                    response.Code = "P001";
                    response.Message = strMessage;
                    response.Success = false;
                    break;
                }
                else
                {
                    response.Code = "000";
                    response.Message = "Success";
                    response.Success = true;
                }

            }
            return response;
        }
        // change flight
        public static CalculateChangeFlightFeesResponse Valid(this IList<Avantik.Web.Service.Message.ManageBooking.ChangeFlight> f, CalculateChangeFlightFeesResponse response, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Change Flight Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowFlightChange == "1")
            {
                foreach (Message.ManageBooking.ChangeFlight cf in f)
                {
                    if(string.IsNullOrEmpty(cf.NewSegment.FlightId))
                    {
                        bError = true;
                        strMessage = "FlightId is required.";
                        code = "P023";
                    }
                    else if (!Helpers.DataType.IsGuid(cf.NewSegment.FlightId))
                    {
                        bError = true;
                        strMessage += "FlightId is invalid.";
                        code = "F011";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.FareId))
                    {
                        bError = true;
                        strMessage = "FareId is required.";
                        code = "P30";
                    }
                    else if (!Helpers.DataType.IsGuid(cf.NewSegment.FareId))
                    {
                        bError = true;
                        strMessage += "FareId is invalid.";
                        code = "F012";
                    }
                    else if (!string.IsNullOrEmpty(cf.NewSegment.TransitFareId) && !Helpers.DataType.IsGuid(cf.NewSegment.TransitFareId))
                    {
                        bError = true;
                        strMessage += "Transit Fare Id is invalid.";
                        code = "F013";
                    }
                    else if (!string.IsNullOrEmpty(cf.NewSegment.TransitFlightId) && !Helpers.DataType.IsGuid(cf.NewSegment.TransitFlightId))
                    {
                        bError = true;
                        strMessage += "Transit Flight Id is invalid.";
                        code = "F014";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.BoardingClassRcd))
                    {
                        bError = true;
                        strMessage = "BoardingClassRcd is required.";
                        code = "P025";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.BoardingClassRcd.Trim()))
                    {
                        bError = true;
                        strMessage = "BoardingClassRcd is required.";
                        code = "P025";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.BookingClassRcd))
                    {
                        bError = true;
                        strMessage = "BookingClassRcd is required.";
                        code = "P031";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.BookingClassRcd.Trim()))
                    {
                        bError = true;
                        strMessage = "BookingClassRcd is required.";
                        code = "P031";
                    }
                    else if (cf.NewSegment.DepartureDate == null || cf.NewSegment.DepartureDate == DateTime.MinValue)
                    {
                        bError = true;
                        strMessage += "DepartureDate is required.";
                        code = "V004";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.OriginRcd))
                    {
                        bError = true;
                        strMessage = "OriginRcd is required.";
                        code = "V002";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.OriginRcd.Trim()))
                    {
                        bError = true;
                        strMessage = "OriginRcd is required.";
                        code = "V002";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.DestinationRcd))
                    {
                        bError = true;
                        strMessage = "DestinationRcd is required.";
                        code = "V003";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.DestinationRcd.Trim()))
                    {
                        bError = true;
                        strMessage = "DestinationRcd is required.";
                        code = "V003";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.OdOriginRcd))
                    {
                        bError = true;
                        strMessage = "OdOriginRcd is required.";
                        code = "V018";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.OdOriginRcd.Trim()))
                    {
                        bError = true;
                        strMessage = "OdOriginRcd is required.";
                        code = "V018";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.OdDestinationRcd))
                    {
                        bError = true;
                        strMessage = "OdDestinationRcd is required.";
                        code = "V019";
                    }
                    else if (string.IsNullOrEmpty(cf.NewSegment.OdDestinationRcd.Trim()))
                    {
                        bError = true;
                        strMessage = "OdDestinationRcd is required.";
                        code = "V019";
                    }
                }

                if (bError)
                {
                    response.Code = code;
                    response.Message = strMessage;
                    response.Success = false;

                }
                else
                {
                    response.Code = "000";
                    response.Message = "Success";
                    response.Success = true;
                }

               
            }
            else
            {
                response.Code = "A008";
                response.Message = "Agency is not allowed to change flight.";
                response.Success = false;
            }

            return response;

        }


        public static BookingCancelResponse Valid(this Avantik.Web.Service.Message.ManageBooking.BookingCancelRequest cancel, BookingCancelResponse response, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Cancel Booking Error: ";

            if (objAuthen.B2bCancelFlight == "1")
            {
                if (bError)
                {
                    response.Code = "";
                    response.Message = strMessage;
                    response.Success = false;
                }
                else
                {
                    response.Success = true;
                }
            }
            else
            {
                response.Code = "P013";
                response.Message = "Agency not allowed to cancel flight.";
                response.Success = false;
            }

            return response;

        }

        public static UpdatedTicketResponse Valid(this IList<Avantik.Web.Service.Message.ManageBooking.UpdatedTicket> objTicket, UpdatedTicketResponse response, Booking booking, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Updated Ticket Error: ";
            string code = string.Empty;

            for (int i = 0; i < objTicket.Count; i++)
            {
                if (booking.IsLockBooking())
                {
                    bError = true;
                    strMessage += "Booking is Locked.";
                    code = "L001";
                }
                else if (string.IsNullOrEmpty(objTicket[i].PassengerID))
                {
                    bError = true;
                    strMessage += "PassengerId is required.";
                    code = "L007";
                }
                else if (!Helpers.DataType.IsGuid(objTicket[i].PassengerID))
                {
                    bError = true;
                    strMessage += "PassengerId is invalid.";
                    code = "L011";
                }
                else if (string.IsNullOrEmpty(objTicket[i].BookingSegmentID))
                {
                    bError = true;
                    strMessage += "BookingSegmentID is required.";
                    code = "F007";
                }
                else if (!Helpers.DataType.IsGuid(objTicket[i].BookingSegmentID))
                {
                    bError = true;
                    strMessage += "BookingSegmentID is invalid.";
                    code = "L012";
                }
                else if (booking.IsPassengerInBooking(objTicket[i].PassengerID.ToString()) == false)
                {
                    bError = true;
                    strMessage += "Passenger is not in booking.";
                    code = "L006";
                }
                else if (booking.IsSegmentInBooking(objTicket[i].BookingSegmentID.ToString()) == false)
                {
                    bError = true;
                    strMessage += "BookingSegmentID is not in booking.";
                    code = "F006";
                }
                else if (!booking.IsFoundTicketNumber(objTicket[i].BookingSegmentID.ToString(), objTicket[i].PassengerID.ToString()))
                {
                    bError = true;
                    strMessage += "Ticket Number not found.";
                    code = "B002";
                }

                // validate segment
                else if (!booking.IsValidSegmentStatus(objTicket[i].BookingSegmentID.ToString()))
                {
                    bError = true;
                    strMessage += "Segment status is invalid.";
                    code = "F005";
                }

                //IsValidFlighttatus
                else if (!booking.IsActiveFlight(new Guid(objTicket[i].BookingSegmentID)))
                {
                    // if flight not active get status to error
                    string[] strError = booking.FlightStatus(new Guid(objTicket[i].BookingSegmentID)).Split(',');
                    bError = true;
                    strMessage += "Flight is " + strError[1] + ".";
                    code = strError[0];
                }
                else if (booking.IsNotValidFlightCheckIn(new Guid(objTicket[i].BookingSegmentID)))
                {
                    bError = true;
                    strMessage += "FlightCheckInStatus is invalid.";
                    code = "F002";
                }
                else if (booking.IsFlownFlight(new Guid(objTicket[i].BookingSegmentID)))
                {
                    bError = true;
                    strMessage += "Flight is flown.";
                    code = "F003";
                }

               // Passenger Check in status 
                else if (booking.IsCheckedPassenger(new Guid(objTicket[i].BookingSegmentID), new Guid(objTicket[i].PassengerID)))
                {
                    bError = true;
                    strMessage += "Passenger status is invalid.";
                    code = "L009";
                }
                else if (booking.IsBoardedPassenger(new Guid(objTicket[i].BookingSegmentID), new Guid(objTicket[i].PassengerID)))
                {
                    bError = true;
                    strMessage += "Passenger status is invalid.";
                    code = "L009";
                }
                else if (booking.IsFlownPassenger(new Guid(objTicket[i].BookingSegmentID), new Guid(objTicket[i].PassengerID)))
                {
                    bError = true;
                    strMessage += "Passenger status is invalid.";
                    code = "L009";
                }
                else if (booking.IsNoShowPassenger(new Guid(objTicket[i].BookingSegmentID), new Guid(objTicket[i].PassengerID)))
                {
                    bError = true;
                    strMessage += "Passenger status is invalid.";
                    code = "L009";
                }
                else if (booking.IsOffLoadedPassenger(new Guid(objTicket[i].BookingSegmentID), new Guid(objTicket[i].PassengerID)))
                {
                    bError = true;
                    strMessage += "Passenger status is invalid.";
                    code = "L009";
                }

                // valid number
                else if (!IsDecimal(objTicket[i].PieceAllowance.ToString()))
                {
                    bError = true;
                    strMessage += "PieceAllowance should be numberic.";
                    code = "B003";
                }
                else if (objTicket[i].PieceAllowance >= 100)
                {
                        bError = true;
                        strMessage += "PieceAllowance should not be equal or greater than 100 units.";
                        code = "B004";
                }
                else if (!IsDecimal(objTicket[i].WeightAllowance.ToString()))
                {
                    bError = true;
                    strMessage += "WeightAllowance should be numberic.";
                    code = "B005";
                }
                else if (objTicket[i].WeightAllowance >= 100000)
                {
                    bError = true;
                    strMessage += "WeightAllowance should not be equal or greater than 100,000 units.";
                    code = "B006";
                }

                // check an error
                if (bError)
                {
                    response.Code = code;
                    response.Message = strMessage;
                    response.Success = false;
                    break;
                }
                else
                {
                    response.Code = "000";
                    response.Message = "Success";
                    response.Success = true;
                }
            }

            return response;
        }

        public static GetFeeResponse Valid(this Avantik.Web.Service.Message.GetFeeDefinitionRequest fee, GetFeeResponse response)
        {
            bool bError = false;
            string strMessage = "GetFeeDefinition Error: ";
            string code = string.Empty;

            if (string.IsNullOrEmpty(fee.FeeRcd))
            {
                bError = true;
                strMessage += "FeeCode is required.";
                code = "V005";
            }
            else if (string.IsNullOrEmpty(fee.FeeRcd.Trim()))
            {
                bError = true;
                strMessage += "FeeCode is required.";
                code = "V005";
            }
            else if (string.IsNullOrEmpty(fee.Origin))
            {
                bError = true;
                strMessage += "OriginRcd is required.";
                code = "V002";
            }

            else if (string.IsNullOrEmpty(fee.Origin.Trim()))
            {
                bError = true;
                strMessage += "OriginRcd is required.";
                code = "V002";
            }
            else if (string.IsNullOrEmpty(fee.Destination))
            {
                bError = true;
                strMessage += "DestinationRcd is required.";
                code = "V003";
            }

            else if (string.IsNullOrEmpty(fee.Destination.Trim()))
            {
                bError = true;
                strMessage += "DestinationRcd is required.";
                code = "V003";
            }
            else if (string.IsNullOrEmpty(fee.Currency))
            {
                bError = true;
                strMessage += "Currency is required.";
                code = "V006";
            }

            else if (string.IsNullOrEmpty(fee.Currency.Trim()))
            {
                bError = true;
                strMessage += "Currency is required.";
                code = "V006";
            }

            if (bError)
            {
                response.Code = code;
                response.Message = strMessage;
                response.Success = false;
            }
            else
            {
                response.Code = "000";
                response.Message = "Success";
                response.Success = true;
            }

            return response;

        }

        public static PaymentFeeResponse Valid(this Avantik.Web.Service.Message.PaymentFeeRequest paymentFee, PaymentFeeResponse response)
        {
            bool bError = false;
            string strMessage = "GetPaymentFee Error: ";
            string code = string.Empty;

            if (string.IsNullOrEmpty(paymentFee.PaymentFee.FormOfPaymentRcd))
            {
                bError = true;
                strMessage += "FormOfPaymentRcd is required.";
                code = "V007";
            }
            else if (string.IsNullOrEmpty(paymentFee.PaymentFee.FormOfPaymentRcd.Trim()))
            {
                bError = true;
                strMessage += "FormOfPaymentRcd is required.";
                code = "V007";
            }
            else if (string.IsNullOrEmpty(paymentFee.PaymentFee.FormOfPaymentSubtypeRcd))
            {
                bError = true;
                strMessage += "FormOfPaymentSubtypeRcd is required.";
                code = "V008";
            }
            else if (string.IsNullOrEmpty(paymentFee.PaymentFee.FormOfPaymentSubtypeRcd.Trim()))
            {
                bError = true;
                strMessage += "FormOfPaymentSubtypeRcd is required.";
                code = "V008";
            }

            //else if (string.IsNullOrEmpty(paymentFee.PaymentFee.Origin.Trim()))
            //{
            //    bError = true;
            //    strMessage += "OriginRcd parameter is required.";
            //}
            //else if (string.IsNullOrEmpty(paymentFee.PaymentFee.Destination.Trim()))
            //{
            //    bError = true;
            //    strMessage += "DestinationRcd parameter is required.";
            //}
            if (bError)
            {
                response.Code = code;
                response.Message = strMessage;
                response.Success = false;
            }
            else
            {
                response.Code = "000";
                response.Message = "Success";
                response.Success = true;
            }

            return response;

        }

        public static AvailabilityResponse Valid(this Avantik.Web.Service.Message.AvailabilityRequest avai, AvailabilityResponse response)
        {
            bool bError = false;
            string strMessage = "GetAvailability Error: ";
            string code = string.Empty;

            if (string.IsNullOrEmpty(avai.OriginRcd))
            {
                bError = true;
                strMessage += "OriginRcd parameter is required.";
                code = "V002";
            }
            if (string.IsNullOrEmpty(avai.OriginRcd.Trim()))
            {
                bError = true;
                strMessage += "OriginRcd parameter is required.";
                code = "V002";
            }
            else if (string.IsNullOrEmpty(avai.DestinationRcd))
            {
                bError = true;
                strMessage += "DestinationRcd parameter is required.";
                code = "V003";
            }

            else if (string.IsNullOrEmpty(avai.DestinationRcd.Trim()))
            {
                bError = true;
                strMessage += "DestinationRcd parameter is required.";
                code = "V003";
            }
            else if (avai.FromDate == null || avai.FromDate == DateTime.MinValue)
            {
                bError = true;
                strMessage += "FromDate parameter is required.";
                code = "V004";
            }
            if (bError)
            {
                response.Code = code;
                response.Message = strMessage;
                response.Success = false;
            }
            else
            {
                response.Code = "000";
                response.Message = "Success";
                response.Success = true;
            }

            return response;

        }

        public static GetSeatMapResponse Valid(this Avantik.Web.Service.Message.SeatMap.GetSeatMapRequest request, GetSeatMapResponse response, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Get SeatMap Error: ";
            string code = string.Empty;

            //if (objAuthen.B2bAllowSeat == "0")
            //{
            //    bError = true;
            //    strMessage += "Agency not allowed seat.";
            //    code = "A008";
            //}
             if (string.IsNullOrEmpty(request.FlightId) || new Guid(request.FlightId) == Guid.Empty)
            {
                bError = true;
                strMessage += "Flight Id is required.";
                code = "P023";
            }
            //else if (string.IsNullOrEmpty(request.BookingClass))
            //{
            //    bError = true;
            //    strMessage += "Booking class is required.";
            //    code = "P025";
            //}

            else if (string.IsNullOrEmpty(request.BoardingClass.Trim()))
            {
                bError = true;
                strMessage += "Boarding class is required.";
                code = "P025";
            }
            if (bError)
            {
                response.Code = code;
                response.Message = strMessage;
                response.Success = false;
            }
            else
            {
                response.Code = "000";
                response.Message = "Success";
                response.Success = true;
            }


            return response;
        }

        public static GetSpecialServicesResponse Valid(this Avantik.Web.Service.Message.GetSpecialServicesRequest request, GetSpecialServicesResponse response, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Get SpecialServices Error: ";
            string code = string.Empty;

            if (objAuthen.B2bAllowSeat == "0")
            {
                bError = true;
                strMessage += "Agency not allowed services.";
                code = "A008";
            }
            else if (string.IsNullOrEmpty(request.ServiceCode))
            {
                bError = true;
                strMessage += "Service code is required.";
                code = "V009";
            }
            else if (string.IsNullOrEmpty(request.ServiceCode.Trim()))
            {
                bError = true;
                strMessage += "Service code is required.";
                code = "V009";
            }

            if (bError)
            {
                response.Code = code;
                response.Message = strMessage;
                response.Success = false;
            }
            else
            {
                response.Code = "000";
                response.Message = "Success";
                response.Success = true;
            }


            return response;
        }

        public static ContactDetailResponse Valid(this Avantik.Web.Service.Message.ContactDetailRequest request, ContactDetailResponse response, Booking booking,Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Updated Contact Error: ";
            string code = string.Empty;

            if (booking.IsLockBooking())
            {
                bError = true;
                strMessage = "Booking is Locked.";
                code = "L001";
            }
            else if (objAuthen.B2bAllowChangeDetail == "0")
            {
                bError = true;
                strMessage = "Agency is not allowed update contactDetail.";
                code = "A008";
            }
            else if (string.IsNullOrEmpty(request.ContactDetail.ContactName) || string.IsNullOrEmpty(request.ContactDetail.ContactName.Trim()))
            {
                bError = true;
                strMessage = "ContactName is required.";
                code = "C002";
            }
            else if (string.IsNullOrEmpty(request.ContactDetail.ContactEmail))
            {
                bError = true;
                strMessage = "ContactEmail is required.";
                code = "C004";
            }
            else if (string.IsNullOrEmpty(request.ContactDetail.ContactEmail.Trim()))
            {
                bError = true;
                strMessage = "ContactEmail is required.";
                code = "C004";
            }
            else if (!IsValidEmail(request.ContactDetail.ContactEmail.Trim()))
            {
                bError = true;
                strMessage = "ContactEmail is invalid.";
                code = "C005";
            }
            else if (string.IsNullOrEmpty(request.ContactDetail.PhoneMobile))
            {
                bError = true;
                strMessage = "PhoneMobile is required.";
                code = "C006";
            }

            else if (string.IsNullOrEmpty(request.ContactDetail.PhoneMobile.Trim()))
            {
                bError = true;
                strMessage = "PhoneMobile is required.";
                code = "C006";
            }

            if(!string.IsNullOrEmpty(request.ContactDetail.MobileEmail))
            {
                if(!IsValidEmail(request.ContactDetail.MobileEmail.Trim()))
                {
                    bError = true;
                    strMessage += "Mobile Email is invalid.";
                    code = "C007";
                }
            }

            if (bError)
            {
                response.Code = code;
                response.Message = strMessage;
                response.Success = false;
            }
            else
            {
                response.Code = "000";
                response.Message = "Success";
                response.Success = true;
            }


            return response;
        }

        public static SegmentCancelResponse Valid(this Avantik.Web.Service.Message.ManageBooking.SegmentCancelRequest seg, SegmentCancelResponse response,Booking booking, Entity.Authentication objAuthen)
        {
            bool bError = false;
            string strMessage = "Cancel Segment Error: ";
            string code = string.Empty;

           // if (objAuthen.B2bAllowSeat == "1")
            {
                    if (booking.IsSegmentInBooking(seg.SegmentId))
                    {
                        bError = true;
                        strMessage += "BookingSegmentId is not in Booking.";
                        code = "F006";
                    }


                    if (bError)
                    {
                        response.Code = code;
                        response.Message = strMessage;
                        response.Success = false;
                    }
                    else
                    {
                        response.Code = "000";
                        response.Message = "Success";
                        response.Success = true;
                    }
            }
            //else
            //{
            //    response.Code = "A008";
            //    response.Message = "Agency is not allowed seat.";
            //    response.Success = false;
            //}

            return response;

        }



        #region Helper
        public static GetSeatMapRequest SeatMapRequest(this GetSeatMapRequest request)
        {
            GetSeatMapRequest result = new GetSeatMapRequest();
            result.BoardingClass = request.BoardingClass.Trim().ToUpper();
            result.BookingClass = request.BookingClass.Trim().ToUpper();
            result.OriginRcd = request.OriginRcd.Trim().ToUpper();
            result.DestinationRcd = request.DestinationRcd.Trim().ToUpper();
            result.FlightId = request.FlightId;
            return result;
        }
        public static GetFeeRequest GetFeeDefinitionMapRequest(this GetFeeDefinitionRequest request, string agencyCode, string currencyRcd, string languageRcd)
        {
            GetFeeRequest req = new GetFeeRequest();
            req.StrAgencyCode = agencyCode.Trim().ToUpper();
            req.StrCurrencyCode = currencyRcd.Trim().ToUpper();
            req.StrLanguage = languageRcd.Trim().ToUpper();
            req.StrClass = request.BookingClass;
            req.StrOrigin = request.Origin.Trim().ToUpper();
            req.StrDestination = request.Destination.Trim().ToUpper();
            req.StrFareBasis = request.FareCode;
            req.StrFeeRcd = request.FeeRcd.Trim().ToUpper();
            req.StrFlightNumber = request.FlightNumber;
            req.DtDate = DateTime.MinValue;
            req.bNoVat = false;

            return req;
        }

        public static IList<Entity.Booking.Fee> StoreEntityFee(this IList<Entity.Booking.Fee> objStoreEntityFee, IList<Entity.Booking.Fee> fees)
        {
            if (objStoreEntityFee != null && fees != null && fees.Count > 0)
            {
                foreach (Entity.Booking.Fee f in fees)
                {
                    objStoreEntityFee.Add(f);
                }
            }
            return objStoreEntityFee;

        }

        public static IList<Entity.Booking.Fee> StoreEntityPaymentVCFee(this IList<Entity.Booking.Fee> objStoreEntityFee,
            IList<Entity.Booking.Fee> fees,
            Guid PassengerId, Guid BookingSegmentId, Guid userId, string agencyCode, string origin, string destination, Guid bookingId
            )
        {
            if (objStoreEntityFee != null && fees != null && fees.Count > 0)
            {

                foreach (Entity.Booking.Fee f in fees)
                {
                    foreach (Entity.Booking.Fee ff in fees)
                    {
                        if (f.FeeId == ff.FeeId)
                        {
                            f.BookingFeeId = Guid.NewGuid();
                            f.AcctFeeAmount = ff.FeeAmount;
                            f.AcctFeeAmountIncl = ff.FeeAmountIncl;
                            f.TotalAmount = ff.FeeAmount;
                            f.TotalAmountIncl = ff.FeeAmountIncl;
                            f.TotalFeeAmount = ff.FeeAmount;
                            f.TotalFeeAmountIncl = ff.FeeAmountIncl;
                         //   f.ChargeAmount = ff.FeeAmount;
                           // f.ChargeAmountIncl = ff.FeeAmountIncl;

                            f.PassengerId = PassengerId;
                            f.BookingSegmentId = BookingSegmentId;
                            f.CreateBy = userId;
                            f.AgencyCode = agencyCode;
                            f.OdOriginRcd = origin;
                            f.OdDestinationRcd = destination;
                            f.NumberOfUnits = 1;
                            f.BookingId = bookingId;
                            // f.agency_currency_rcd = 
                        }
                    }

                }



                foreach (Entity.Booking.Fee f in fees)
                {
                    objStoreEntityFee.Add(f);
                }
            }
            return objStoreEntityFee;

        }


        public static IList<Entity.Booking.Fee> StoreEntityMessageFee(this IList<Message.ManageBooking.Fee> objMessageFee,
            IList<Entity.Booking.Fee> objStoreEntityFee,Guid userId, string agency)
        {
            if (objMessageFee != null && objStoreEntityFee != null && objMessageFee.Count > 0)
            {
                foreach (Entity.Booking.Fee f in objStoreEntityFee)
                {
                    foreach (Message.ManageBooking.Fee ff in objMessageFee)
                    {
                        f.FeeId = Guid.NewGuid();
                        f.FeeRcd = ff.FeeRcd;
                        f.BookingFeeId = Guid.NewGuid();
                        f.AcctFeeAmount = ff.FeeAmount;
                        f.AcctFeeAmountIncl =ff.FeeAmountIncl;
                        f.TotalAmount = ff.FeeAmount;
                        f.TotalAmountIncl = ff.FeeAmountIncl;
                        f.TotalFeeAmount = ff.FeeAmount;
                        f.TotalFeeAmountIncl = ff.FeeAmountIncl;
                        f.ChargeAmount = ff.FeeAmount;
                        f.ChargeAmountIncl = ff.FeeAmountIncl;

                        f.PassengerId = ff.PassengerId;
                        f.BookingSegmentId = ff.BookingSegmentId;
                        f.CreateBy = userId;
                        f.AgencyCode = agency;
                        f.OdOriginRcd = ff.OriginRcd;
                        f.OdDestinationRcd = ff.DestinationRcd;
                        f.NumberOfUnits = 0;
                        f.BookingId = ff.BookingId;
                        f.Comment = ff.Comment;
                    }

                }



                foreach (Entity.Booking.Fee f in objStoreEntityFee)
                {
                    objStoreEntityFee.Add(f);
                }
            }

            return objStoreEntityFee;
        }


        public static bool IsAllDigits(string s)
        {
            return s.All(Char.IsDigit);
        }

        private static bool IsInt(string strInt)
        {
            bool result;
            int Num;
            bool isNum = Int32.TryParse(strInt.ToString(), out Num);

            if (isNum)
                result = true; //integer
            else
                result = false;

            return result;
        }
        private static bool IsDecimal(string strNum)
        {
            bool result;
            decimal Num;
            bool isNum = decimal.TryParse(strNum, out Num);

            if (isNum)
                result = true; //decimal
            else
                result = false;

            return result;
        }

        public static bool IsValidEmail(string emailaddress)
        {
            try
            {
                MailAddress m = new MailAddress(emailaddress);

                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        public static bool IsDuplicateSeat(IList<Avantik.Web.Service.Message.ManageBooking.SeatAssign> seats, string SeatNumber)
        {
            bool result = false;
            for (int j = 1; j < seats.Count; j++)
            {
                //if (seats[j].SeatNumber.ToUpper().Equals(SeatNumber))
                //{
                //    result = true;
                //    break;
                //}
            }
            return result;
        }

        public static void SetDefaultTransitFlight(this IList<Avantik.Web.Service.Message.ManageBooking.ChangeFlight> f)
        {
            // fill to protect error
            foreach (Message.ManageBooking.ChangeFlight cf in f)
            {

                if (string.IsNullOrEmpty(cf.NewSegment.TransitFlightId))
                {
                    cf.NewSegment.TransitFlightId = "00000000-0000-0000-0000-000000000000";
                }
                if (string.IsNullOrEmpty(cf.NewSegment.TransitFareId))
                {
                    cf.NewSegment.TransitFareId = "00000000-0000-0000-0000-000000000000";
                }
            }

        }


        #endregion


    }
}
