using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Avantik.Web.Service.Model.Contract;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.COMHelper;
using entity = Avantik.Web.Service.Entity.Booking;
using System.Data;
using ADODB;
using Avantik.Web.Service.Model.COM.Extension;
using Avantik.Web.Service.Exception.Booking;
using Avantik.Web.Service.Entity.Agency;
using System.Collections;
using System.Data.OleDb;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Entity.Booking;
using System.Data.SqlClient;

namespace Avantik.Web.Service.Model.COM
{
    public class BookingModelService : RunComplus ,IBookingModelService
    {
        string _server = string.Empty;
        string _user = string.Empty;
        string _pass = string.Empty;
        string _domain = string.Empty;

        bool result = false;
        public BookingModelService(string server, string user, string pass, string domain)
            :base(user,pass,domain)
        {
            _server = server;
            _user = user;
            _pass = pass;
            _domain = domain;
        }
        public bool SaveBooking(entity.Booking booking,
                                bool createTickets,
                                bool readBooking,
                                bool readOnly,
                                bool bSetLock,
                                bool bCheckSeatAssignment,
                                bool bCheckSessionTimeOut)
        {
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsFlight = null;
            ADODB.Recordset rsQuote = null;

            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strUserId = string.Empty;
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;

            short iAdult = 0;
            short iChild = 0;
            short iInfant = 0;
            short iOther = 0;

            tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                objBooking.GetEmpty(ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment,
                                    ref rsMapping,
                                    ref rsService,
                                    ref rsTax,
                                    ref rsFees,
                                    ref rsFlight,
                                    ref rsQuote,
                                    ref strBookingId,
                                    ref strAgencyCode,
                                    ref strCurrency,
                                    ref iAdult,
                                    ref iChild,
                                    ref iInfant,
                                    ref iOther,
                                    ref strOther,
                                    ref strUserId);

                if (booking.Header != null)
                {
                    booking.Header.ToRecordset(ref rsHeader);
                }
                if (booking.Segments != null && booking.Segments.Count > 0)
                {
                    booking.Segments.ToRecordset(ref rsSegment);
                }
                if (booking.Passengers != null && booking.Passengers.Count > 0)
                {
                    booking.Passengers.ToRecordset(ref rsPassenger);
                }
                if (booking.Fees != null && booking.Fees.Count > 0)
                {
                    booking.Fees.ToRecordset(ref rsFees);
                }
                if (booking.Remarks != null && booking.Remarks.Count > 0)
                {
                    booking.Remarks.ToRecordset(ref rsRemark);
                }
                if (booking.Payments != null && booking.Payments.Count > 0)
                {
                    booking.Payments.ToRecordset(ref rsPayment);
                }
                if (booking.Mappings != null && booking.Mappings.Count > 0)
                {
                    booking.Mappings.ToRecordset(ref rsMapping);
                }
                if (booking.Services != null && booking.Services.Count > 0)
                {
                    booking.Services.ToRecordset(ref rsService);
                }
                if (booking.Taxs != null && booking.Taxs.Count > 0)
                {
                    booking.Taxs.ToRecordset(ref rsTax);
                }

                enumSaveError = objBooking.Save(ref rsHeader,
                                                  ref rsSegment,
                                                  ref rsPassenger,
                                                  ref rsRemark,
                                                  ref rsPayment,
                                                  ref rsMapping,
                                                  ref rsService,
                                                  ref rsTax,
                                                  ref rsFees,
                                                  ref createTickets,
                                                  ref readBooking,
                                                  ref readOnly,
                                                  ref bSetLock,
                                                  ref bCheckSeatAssignment,
                                                  ref bCheckSessionTimeOut);

                if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                {
                    result = true;
                }
                else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                {
                    throw new BookingSaveException("BOOKINGINUSE");
                }
                else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                {
                    throw new BookingSaveException("BOOKINGREADERROR");
                }
                else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                {
                    throw new BookingSaveException("DATABASEACCESS");
                }
                else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                {
                    throw new BookingSaveException("DUPLICATESEAT","277");
                }
                else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                {
                    throw new BookingSaveException("SESSIONTIMEOUT", "H005");
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }

            return result;
        }

        public entity.Booking ReadBooking(string bookingId,
                                   string bookingReference,
                                   double bookingNumber,
                                   bool bReadonly,
                                   bool bSeatLock,
                                   string userId,
                                   bool bReadHeader,
                                   bool bReadSegment,
                                   bool bReadPassenger,
                                   bool bReadRemark,
                                   bool bReadPayment,
                                   bool bReadMapping,
                                   bool bReadService,
                                   bool bReadTax,
                                   bool bReadFee,
                                   bool bReadOd,
                                   string releaseBookingId,
                                   string CompleteRemarkId)

        {
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;

            entity.Booking booking = new entity.Booking();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                if (!String.IsNullOrEmpty(bookingId))
                {
                    if (bookingId.Equals(Guid.Empty.ToString()))
                        bookingId = string.Empty;

                    if (objBooking.Read("{" + bookingId.ToUpper() + "}",
                                        bookingReference,
                                        bookingNumber,
                                        ref rsHeader,
                                        ref rsSegment,
                                        ref rsPassenger,
                                        ref rsRemark,
                                        ref rsPayment,
                                        ref rsMapping,
                                        ref rsService,
                                        ref rsTax,
                                        ref rsFees,
                                        ref rsQuote,
                                        ref bReadonly,
                                        ref bSeatLock,
                                        ref userId,
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
                                        ref releaseBookingId,
                                        ref CompleteRemarkId) == true)

                    {

                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);
                        }
                        if (rsSegment != null && rsSegment.RecordCount > 0)
                        {
                            booking.Segments = booking.Segments.FillBooking(rsSegment);
                        }
                        if (rsPassenger != null && rsPassenger.RecordCount > 0)
                        {
                            booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
                        }
                        if (rsFees != null && rsFees.RecordCount > 0)
                        {
                            booking.Fees = booking.Fees.FillBooking(rsFees);
                        }
                        if (rsRemark != null && rsRemark.RecordCount > 0)
                        {
                            booking.Remarks = booking.Remarks.FillBooking(rsRemark);
                        }
                        if (rsQuote != null && rsQuote.RecordCount > 0)
                        {
                            booking.Quotes = booking.Quotes.FillBooking(rsQuote);
                        }
                        if (rsTax != null && rsTax.RecordCount > 0)
                        {
                            booking.Taxs = booking.Taxs.FillBooking(rsTax);
                        }
                        if (rsService != null && rsService.RecordCount > 0)
                        {
                            booking.Services = booking.Services.FillBooking(rsService);
                        }
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        }
                        if (rsPayment != null && rsPayment.RecordCount > 0)
                        {
                            booking.Payments = booking.Payments.FillBooking(rsPayment);
                        }
                    }
                }
            }
            catch (BookingException bookingExc)
            {
                throw bookingExc;
            }
            catch
            {
                throw new System.Exception("Read booking is error");
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }

            return booking;
        }

        public bool CancelBooking(string bookingId,
                                  string bookingReference,
                                  double bookingNumber,
                                  string userId,
                                  string agencyCode,
                                  bool bWaveFee,
                                  bool bVoid)
        {
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            entity.Booking booking = new entity.Booking();
            bool bCancelBooking = false;
            bool bCancelSegment = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                if (!String.IsNullOrEmpty(bookingId) && !String.IsNullOrEmpty(userId))
                {
                    Agent agent = new Agent();
                    AgencyService agencyService = new AgencyService(_server, _user, _pass, _domain);
                    agent = agencyService.GetAgencySessionProfile(agencyCode, userId);

                    if (objBooking.Read("{" + bookingId.ToUpper() + "}",
                                        bookingReference,
                                        bookingNumber,
                                        ref rsHeader,
                                        ref rsSegment,
                                        ref rsPassenger,
                                        ref rsRemark,
                                        ref rsPayment,
                                        ref rsMapping,
                                        ref rsService,
                                        ref rsTax,
                                        ref rsFees,
                                        ref rsQuote,
                                        false,
                                        false,
                                        "{" + userId.ToUpper() + "}",
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        false,
                                        true,
                                        string.Empty,
                                        string.Empty) == true)
                    {
                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);
                        }
                        if (rsSegment != null && rsSegment.RecordCount > 0)
                        {
                            booking.Segments = booking.Segments.FillBooking(rsSegment);
                        }
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        }


                        if (!booking.Header.BookingId.Equals(Guid.Empty) && !booking.Header.BookingId.Equals(String.Empty))
                        {
                            if (!booking.IsLockBooking())
                            {
                                bool processTicketRefund = false;
                                bool IncludeRefundableOnly = true;
                                DateTime dtNoShowSegmentDatetime = new DateTime();
                                bool bReturnConfirmCancelStatus = true;
                                Int32 cntBoarded = 0;
                                Int32 cntChecked = 0;
                                Int32 cntFlown = 0;

                                //Get ProcessRefundFlag for Check Agency Refund
                                if (agent.ProcessRefundFlag == 1)
                                    processTicketRefund = true;

                                for (int i = 0; i < booking.Segments.Count; i++)
                                {
                                    if (!booking.IsCheckedPassenger(booking.Segments[i].BookingSegmentId))
                                    {
                                        if (!booking.IsBoardedPassenger(booking.Segments[i].BookingSegmentId))
                                        {
                                            if (!booking.IsFlownPassenger(booking.Segments[i].BookingSegmentId))
                                            {
                                                bCancelSegment = objBooking.SegmentCancel("{" + Convert.ToString(booking.Segments[i].BookingSegmentId).ToUpper() + "}",
                                                                                ref rsSegment,
                                                                                ref rsMapping,
                                                                                ref rsService,
                                                                                ref rsPayment,
                                                                                ref rsTax,
                                                                                ref rsQuote,
                                                                                ref userId,
                                                                                ref agencyCode,
                                                                                ref bWaveFee,
                                                                                ref processTicketRefund,
                                                                                ref IncludeRefundableOnly,
                                                                                ref dtNoShowSegmentDatetime,
                                                                                ref bVoid,
                                                                                ref bReturnConfirmCancelStatus);
                                                if (bCancelSegment != true) return false;
                                            }
                                            else
                                            {
                                                cntFlown++;
                                                if (cntFlown == booking.Segments.Count)
                                                    throw new ModifyBookingException("Passenger check in status as FLOWN", "L005");
                                            }
                                        }
                                        else
                                        {
                                            cntBoarded++;
                                            if (cntBoarded == booking.Segments.Count)
                                                throw new ModifyBookingException("Passenger check in status as BOARDED", "L004");
                                        }
                                    }
                                    else
                                    {
                                        cntChecked++;
                                        if (cntChecked == booking.Segments.Count)
                                            throw new ModifyBookingException("Passenger check in status as CHECKED", "L003");
                                    }
                                }

                                tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;

                                if (bCancelSegment)
                                {
                                    enumSaveError = objBooking.Save(ref rsHeader,
                                                                     ref rsSegment,
                                                                     ref rsPassenger,
                                                                     ref rsRemark,
                                                                     ref rsPayment,
                                                                     ref rsMapping,
                                                                     ref rsService,
                                                                     ref rsTax,
                                                                     ref rsFees,
                                                                     false,
                                                                     false,
                                                                     false,
                                                                     false,
                                                                     false,
                                                                     false);
                                }
                                if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                                {
                                    bCancelBooking = true;
                                }
                                else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                                {
                                    throw new BookingSaveException("BOOKINGINUSE");
                                }
                                else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                                {
                                    throw new BookingSaveException("BOOKINGREADERROR");
                                }
                                else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                                {
                                    throw new BookingSaveException("DATABASEACCESS");
                                }
                                else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                                {
                                    throw new BookingSaveException("DUPLICATESEAT");
                                }
                                else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                                {
                                    throw new BookingSaveException("SESSIONTIMEOUT");
                                }
                            }
                            else
                            {
                                throw new ModifyBookingException("Booking is Locked.", "L001");
                            }
                        }

                    }
                }
            }
            catch (BookingSaveException saveExc)
            {
                throw saveExc;
            }
            catch (ModifyBookingException bookingExc)
            {
                throw bookingExc;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }

            return bCancelBooking;
        }

        public bool CancelSegment(string bookingId,
                                  string segmentId,
                                  string userId,
                                  string agencyCode,
                                  bool bWaveFee,
                                  bool bVoid)
        {
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            entity.Booking booking = new entity.Booking();
            bool bCancelSegment = false;
            bool IsSave = false;
            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                if (!String.IsNullOrEmpty(bookingId) && !bookingId.Equals(Guid.Empty.ToString()))
                {
                    //Agent agent = new Agent();
                    //AgencyService agencyService = new AgencyService(_server, _user, _pass, _domain);
                    //agent = agencyService.GetAgencySessionProfile(agencyCode, userId);

                   // if (agent != null && agent.B2BAllowCancelFlight == 1)
                    {
                        if (objBooking.Read("{" + bookingId.ToUpper() + "}",
                                            string.Empty,
                                            0,
                                            ref rsHeader,
                                            ref rsSegment,
                                            ref rsPassenger,
                                            ref rsRemark,
                                            ref rsPayment,
                                            ref rsMapping,
                                            ref rsService,
                                            ref rsTax,
                                            ref rsFees,
                                            ref rsQuote,
                                            false,
                                            false,
                                            "{" + userId.ToUpper() + "}",
                                            true,
                                            true,
                                            true,
                                            true,
                                            true,
                                            true,
                                            true,
                                            true,
                                            false,
                                            true,
                                            string.Empty,
                                            string.Empty) == true)
                        {
                            if (rsHeader != null && rsHeader.RecordCount > 0)
                            {
                                booking.Header = booking.Header.FillBooking(rsHeader);
                            }
                            if (rsSegment != null && rsSegment.RecordCount > 0)
                            {
                                booking.Segments = booking.Segments.FillBooking(rsSegment);
                            }
                            if (rsMapping != null && rsMapping.RecordCount > 0)
                            {
                                booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                            }
                            if (rsPassenger != null && rsPassenger.RecordCount > 0)
                            {
                                booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
                            }

                            if (!booking.Header.BookingId.Equals(Guid.Empty) && !booking.Header.BookingId.Equals(String.Empty))
                            {
                                if (!booking.IsLockBooking())
                                {
                                    bool processTicketRefund = false;
                                    bool IncludeRefundableOnly = true;
                                    DateTime dtNoShowSegmentDatetime = new DateTime();
                                    bool bReturnConfirmCancelStatus = true;

                                    //Get ProcessRefundFlag for Check Agency Refund
                                   // if (agent.ProcessRefundFlag == 1)
                                        processTicketRefund = true;

                                    if (rsSegment != null && rsSegment.RecordCount > 0)
                                    {
                                        Guid flightConnectionId = Guid.Empty;
                                        Guid tmpSegmentId = Guid.Empty;

                                        for (int i = 0; i < booking.Segments.Count; i++)
                                        {
                                            if (!booking.IsCheckedPassenger(booking.Segments[i].BookingSegmentId))
                                            {
                                                if (!booking.IsFlownPassenger(booking.Segments[i].BookingSegmentId))
                                                {
                                                    if (!booking.IsBoardedPassenger(booking.Segments[i].BookingSegmentId))
                                                    {
                                                        if (booking.IsSegmentInBooking(segmentId))
                                                        {
                                                            if (booking.IsValidSegmentStatus(segmentId))
                                                            {
                                                                if (booking.Segments[i].BookingSegmentId == Guid.Parse(segmentId))
                                                                {
                                                                    flightConnectionId = booking.Segments[i].FlightConnectionId;
                                                                    tmpSegmentId = Guid.Parse(segmentId);

                                                                    bCancelSegment = objBooking.SegmentCancel("{" + booking.Segments[i].BookingSegmentId + "}",
                                                                                                    ref rsSegment,
                                                                                                    ref rsMapping,
                                                                                                    ref rsService,
                                                                                                    ref rsPayment,
                                                                                                    ref rsTax,
                                                                                                    ref rsQuote,
                                                                                                    ref userId,
                                                                                                    ref agencyCode,
                                                                                                    ref bWaveFee,
                                                                                                    ref processTicketRefund,
                                                                                                    ref IncludeRefundableOnly,
                                                                                                    ref dtNoShowSegmentDatetime,
                                                                                                    ref bVoid,
                                                                                                    ref bReturnConfirmCancelStatus);


                                                                    if (bCancelSegment)
                                                                    {
                                                                        IsSave = true;
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                throw new ModifyBookingException("SegmentId already cancelled.", "302");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            throw new ModifyBookingException("BookingSegmentId is not in Booking.", "F006");
                                                        }

                                                    }
                                                    else
                                                    {
                                                        throw new ModifyBookingException("Passenger check in status as BOARDED", "L004");
                                                    }
                                                }
                                                else
                                                {
                                                    throw new ModifyBookingException("Passenger check in status as FLOWN", "L005");
                                                }
                                            }
                                            else
                                            {
                                                throw new ModifyBookingException("Passenger check in status as CHECKED", "L003");
                                            }
                                        }
                                        // connecting flight  need to review
                                        if (flightConnectionId != Guid.Empty)
                                        {
                                            for (int i = 0; i < booking.Segments.Count; i++)
                                            {
                                                if (booking.Segments[i].FlightConnectionId == flightConnectionId && booking.Segments[i].BookingSegmentId != tmpSegmentId)
                                                {
                                                    bCancelSegment = objBooking.SegmentCancel("{" + booking.Segments[i].BookingSegmentId + "}",
                                                                                    ref rsSegment,
                                                                                    ref rsMapping,
                                                                                    ref rsService,
                                                                                    ref rsPayment,
                                                                                    ref rsTax,
                                                                                    ref rsQuote,
                                                                                    ref userId,
                                                                                    ref agencyCode,
                                                                                    ref bWaveFee,
                                                                                    ref processTicketRefund,
                                                                                    ref IncludeRefundableOnly,
                                                                                    ref dtNoShowSegmentDatetime,
                                                                                    ref bVoid,
                                                                                    ref bReturnConfirmCancelStatus);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    throw new ModifyBookingException("Booking is locked.", "L001");
                                }
                            }

                            tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;

                            if (IsSave)
                            {
                                enumSaveError = objBooking.Save(ref rsHeader,
                                                                 ref rsSegment,
                                                                 ref rsPassenger,
                                                                 ref rsRemark,
                                                                 ref rsPayment,
                                                                 ref rsMapping,
                                                                 ref rsService,
                                                                 ref rsTax,
                                                                 ref rsFees,
                                                                 false,
                                                                 false,
                                                                 false,
                                                                 false,
                                                                 false,
                                                                 false);
                            }
                            if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                            {
                                bCancelSegment = true;
                            }
                            else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                            {
                                throw new BookingSaveException("BOOKINGINUSE");
                            }
                            else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                            {
                                throw new BookingSaveException("BOOKINGREADERROR");
                            }
                            else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                            {
                                throw new BookingSaveException("DATABASEACCESS");
                            }
                            else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                            {
                                throw new BookingSaveException("DUPLICATESEAT");
                            }
                            else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                            {
                                throw new BookingSaveException("SESSIONTIMEOUT");
                            }
                        }
                    }
                    //else
                    //{
                    //    throw new ModifyBookingException("Agents not allowed cancel flight", "P014");
                    //}
                }
                else
                {
                    throw new ModifyBookingException("Require Booking ID", "B009");
                }
            }
            catch (BookingSaveException saveExc)
            {
                throw saveExc;
            }
            catch (ModifyBookingException bookingExc)
            {
                throw bookingExc;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }

            return bCancelSegment;
        }

        public entity.Booking BookFlight(string strAgencyCode,
                                            string strCurrency,
                                            IList<entity.Flight> flight,
                                            string strBookingId,
                                            short iAdult,
                                            short iChild,
                                            short iInfant,
                                            short iOther,
                                            string strOther,
                                            string strUserId,
                                            string strIpAddress,
                                            string strLanguageCode,
                                            bool bNoVat)
        {
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            ADODB.Recordset rsFlights = RecordsetHelper.FabricateFlightRecordset();

            entity.Booking booking = new entity.Booking();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                if (flight != null && flight.Count > 0)
                {
                    flight.ToRecordset(ref rsFlights);
                }

               string result = objBooking.GetEmpty(ref rsHeader,
                                ref rsSegment,
                                ref rsPassenger,
                                ref rsRemark,
                                ref rsPayment,
                                ref rsMapping,
                                ref rsService,
                                ref rsTax,
                                ref rsFees,
                                ref rsFlights,
                                ref rsQuote,
                                ref strBookingId,
                                ref strAgencyCode,
                                ref strCurrency,
                                ref iAdult,
                                ref iChild,
                                ref iInfant,
                                ref iOther,
                                ref strOther,
                                ref strUserId,
                                ref strIpAddress,
                                ref strLanguageCode,
                                ref bNoVat) ;

               if (result != string.Empty)
               {
                   if (rsHeader != null && rsHeader.RecordCount > 0)
                   {
                       booking.Header = booking.Header.FillBooking(rsHeader);
                   }
                   if (rsSegment != null && rsSegment.RecordCount > 0)
                   {
                       booking.Segments = booking.Segments.FillBooking(rsSegment);
                   }
                   if (rsPassenger != null && rsPassenger.RecordCount > 0)
                   {
                       booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
                   }
                   if (rsFees != null && rsFees.RecordCount > 0)
                   {
                       booking.Fees = booking.Fees.FillBooking(rsFees);
                   }
                   if (rsRemark != null && rsRemark.RecordCount > 0)
                   {
                       booking.Remarks = booking.Remarks.FillBooking(rsRemark);
                   }
                   if (rsQuote != null && rsQuote.RecordCount > 0)
                   {
                       booking.Quotes = booking.Quotes.FillBooking(rsQuote);
                   }
                   if (rsTax != null && rsTax.RecordCount > 0)
                   {
                       booking.Taxs = booking.Taxs.FillBooking(rsTax);
                   }
                   if (rsService != null && rsService.RecordCount > 0)
                   {
                       booking.Services = booking.Services.FillBooking(rsService);
                   }
                   if (rsMapping != null && rsMapping.RecordCount > 0)
                   {
                       booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                   }
                   if (rsPayment != null && rsPayment.RecordCount > 0)
                   {
                       booking.Payments = booking.Payments.FillBooking(rsPayment);
                   }
               }
               else
               {
                   throw new BookingException("Add flight fail : " + result);
               }
            }
            catch
            {
                throw ;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }
            return booking;
        }

        public entity.Booking CalculateBookingChange(string strBookingId,
                                            IList<entity.Flight> flight,
                                            string strUserId,
                                            string strAgencyCode,
                                            bool bNoVat)
        {
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            ADODB.Recordset rsFlights = RecordsetHelper.FabricateFlightRecordset();

            entity.Booking booking = new entity.Booking();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                if (!String.IsNullOrEmpty(strBookingId))
                {
                    Boolean resultAdd = false;

                    resultAdd = BookingChangeFlight(strBookingId,
                                                    ref rsHeader,
                                                    ref rsSegment,
                                                    ref rsPassenger,
                                                    ref rsRemark,
                                                    ref rsPayment,
                                                    ref rsMapping,
                                                    ref rsService,
                                                    ref rsTax,
                                                    ref rsFees,
                                                    ref rsQuote,
                                                    flight,
                                                    strUserId,
                                                    strAgencyCode,
                                                    "",
                                                    "",
                                                    bNoVat);
                    if (resultAdd)
                    {
                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);
                        }
                        if (rsSegment != null && rsSegment.RecordCount > 0)
                        {
                            booking.Segments = booking.Segments.FillBooking(rsSegment);
                        }
                        if (rsPassenger != null && rsPassenger.RecordCount > 0)
                        {
                            booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
                        }
                        if (rsFees != null && rsFees.RecordCount > 0)
                        {
                            booking.Fees = booking.Fees.FillBooking(rsFees);
                        }
                        if (rsRemark != null && rsRemark.RecordCount > 0)
                        {
                            booking.Remarks = booking.Remarks.FillBooking(rsRemark);
                        }
                        if (rsQuote != null && rsQuote.RecordCount > 0)
                        {
                            booking.Quotes = booking.Quotes.FillBooking(rsQuote);
                        }
                        if (rsTax != null && rsTax.RecordCount > 0)
                        {
                            booking.Taxs = booking.Taxs.FillBooking(rsTax);
                        }
                        if (rsService != null && rsService.RecordCount > 0)
                        {
                            booking.Services = booking.Services.FillBooking(rsService);
                        }
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        }
                        if (rsPayment != null && rsPayment.RecordCount > 0)
                        {
                            booking.Payments = booking.Payments.FillBooking(rsPayment);
                        }

                    }else
                    {
                        booking = null;
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }
            return booking;
        }
        
        public entity.Booking ChangeFlightFee(string strBookingId,
                                    IList<entity.Flight> flight,
                                    string strUserId,
                                    string agencyCode,string currency,string language,
                                    bool bNoVat)
        {
            tikAeroProcess.Booking objBooking = null;
            tikAeroProcess.clsCreditCard objCreditCard = null;
            tikAeroProcess.Payment objPayment = null;
            tikAeroProcess.Fees objFees = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            ADODB.Recordset rsFlights = RecordsetHelper.FabricateFlightRecordset();

            string strXml = string.Empty;
            Boolean resultAddFlight = false;
            string paymentType = string.Empty;
            entity.Booking bookingResponse = new entity.Booking();
            entity.Booking booking = new entity.Booking();
            var objSegmentIdMapping = new List<KeyValuePair<string, string>>();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;

                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                    objFees = new tikAeroProcess.Fees();
                }

                if (!String.IsNullOrEmpty(strBookingId))
                {
                    //read booking first for prepare RS 
                    if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                        string.Empty,
                        0,
                        ref rsHeader,
                        ref rsSegment,
                        ref rsPassenger,
                        ref rsRemark,
                        ref rsPayment,
                        ref rsMapping,
                        ref rsService,
                        ref rsTax,
                        ref rsFees,
                        ref rsQuote,
                        false,
                        false,
                        "{" + strUserId.ToUpper() + "}",
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        string.Empty,
                        string.Empty) == true)
                    {
                        // convert rs to booking for preparing header and mapping  value
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        }
                    }


                    // Found flight need to add flight first  return rs fee,mapping,segment
                    if (flight != null && flight.Count > 0)
                    {
                        // we can get rs from add flight
                        resultAddFlight = BookingChangeFlight(strBookingId,
                                                        ref rsHeader,
                                                        ref rsSegment,
                                                        ref rsPassenger,
                                                        ref rsRemark,
                                                        ref rsPayment,
                                                        ref rsMapping,
                                                        ref rsService,
                                                        ref rsTax,
                                                        ref rsFees,
                                                        ref rsQuote,
                                                        flight,
                                                        strUserId,
                                                        agencyCode,
                                                        currency,
                                                        language,
                                                        bNoVat);
                    }
                        
                    //success add flight return fee,seg,mapping
                    if (resultAddFlight)
                    {
                        // if found US status  return error
                        if (objSegmentIdMapping == null || objSegmentIdMapping.Count == 0)
                        {
                        }

                        if (rsFees != null && rsFees.RecordCount > 0)
                        {
                            booking.Fees = booking.Fees.FillBooking(rsFees);
                        }

                        bookingResponse.Fees = booking.Fees;
                    }
                    else
                    {
                        // add flight fail
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                if (objCreditCard != null)
                {
                    Marshal.FinalReleaseComObject(objCreditCard);
                    objCreditCard = null;
                }
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }
            return bookingResponse;
        }


        public entity.Booking  ModifyBooking(string strBookingId,
                                            IList<entity.Flight> flight,
                                            IList<entity.PassengerService> services,
                                            IList<Entity.SeatAssign> seatAssigns,
                                            IList<Entity.Booking.Fee> fees,
                                            IList<Entity.NameChange> NameChange,
                                            IList<entity.Payment> payments,
                                            string strUserId,
                                            string agencyCode,
                                            string currencyRcd,
                                            string languageRcd,
                                            bool bNoVat,
                                            string actionCode)
        {
            tikAeroProcess.Booking objBooking = null;
           // tikAeroProcess.clsCreditCard objCreditCard = null;
            tikAeroProcess.Payment objPayment = null;
            tikAeroProcess.Fees objFees = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            ADODB.Recordset rsFlights = RecordsetHelper.FabricateFlightRecordset();

            string strXml = string.Empty;
            Boolean resultAddFlight = false;
            bool createTickets = false;
            bool readBooking = false;
            bool readOnly = false;
            bool bSetLock = false;
            // can not check dup seat in save booking because we save payment already
            bool bCheckSeatAssignment = false;
            bool bCheckSessionTimeOut = false;
            string paymentType = string.Empty;
            entity.Booking bookingResponse = new entity.Booking();
            entity.Booking booking = new entity.Booking();
            var objSegmentIdMapping = new List<KeyValuePair<string, string>>();
            bool IsProcessSuccess = true;
            bool IsPaymentSuccess = false;
            decimal totalOutStanding = 0;
            string strError = string.Empty;
            tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                    objFees = new tikAeroProcess.Fees();
                    objPayment = new tikAeroProcess.Payment();
                }

                if (!String.IsNullOrEmpty(strBookingId))
                {
                    //read booking first for prepare RS 
                    if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                        string.Empty,
                        0,
                        ref rsHeader,
                        ref rsSegment,
                        ref rsPassenger,
                        ref rsRemark,
                        ref rsPayment,
                        ref rsMapping,
                        ref rsService,
                        ref rsTax,
                        ref rsFees,
                        ref rsQuote,
                        false,
                        false,
                        "{" + strUserId.ToUpper() + "}",
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        string.Empty,
                        string.Empty) == true)
                    {
                        // convert rs to booking for preparing header and mapping  value
                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);
                        }
                        if (rsPassenger != null && rsPassenger.RecordCount > 0)
                        {
                            booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
                        }
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        }
                        if (rsRemark != null && rsRemark.RecordCount > 0)
                        {
                            booking.Remarks = booking.Remarks.FillBooking(rsRemark);
                        }
                        if (rsFees != null && rsFees.RecordCount > 0)
                        {
                            booking.Fees = booking.Fees.FillBooking(rsFees);
                        }

                        // Found flight need to add flight first  return RS of fee,mapping,segment
                        if (flight != null && flight.Count > 0)
                        {
                            // check infant over limit
                            string InfantOverLimitFlag = ConfigHelper.ToString("InfantOverLimitFlag");
                            if (InfantOverLimitFlag.ToUpper() == "TRUE")
                            {
                                int numberOfInfant = 0;
                                // count inf
                                foreach (Entity.Booking.Passenger p in booking.Passengers)
                                {
                                    if (p.PassengerTypeRcd == "INF")
                                    {
                                        numberOfInfant += 1;
                                    }
                                }

                                for (int i = 0; i < flight.Count; i++)
                                {
                                    bool result = InfantOverLimit(numberOfInfant, flight[i].FlightId.ToString(), flight[i].OdOriginRcd, flight[i].DestinationRcd, flight[i].BoardingClassRcd);

                                    if (result)
                                    {
                                        IsProcessSuccess = false;
                                        throw new ModifyBookingException("Add flight fail: Over Infant", "F011");
                                    }
                                }
                            }

                            // get rs from add flight
                            resultAddFlight = BookingChangeFlight(strBookingId,
                                                            ref rsHeader,
                                                            ref rsSegment,
                                                            ref rsPassenger,
                                                            ref rsRemark,
                                                            ref rsPayment,
                                                            ref rsMapping,
                                                            ref rsService,
                                                            ref rsTax,
                                                            ref rsFees,
                                                            ref rsQuote,
                                                            flight,
                                                            strUserId,
                                                            agencyCode,
                                                            currencyRcd,
                                                            languageRcd,
                                                            bNoVat);
                            //success add flight return fee,seg,mapping
                            if (resultAddFlight)
                            {
                                //making segment id mapping : NOT yet suport connection flight
                                if (rsSegment != null && rsSegment.RecordCount > 0)
                                {
                                    booking.Segments = booking.Segments.FillBooking(rsSegment);
                                }

                                // find new segment: key == old seg, Value = new Seg
                                objSegmentIdMapping = FindNewSegmentFlightChange(booking.Segments);

                                // if found US status  return error
                                if (objSegmentIdMapping == null || objSegmentIdMapping.Count == 0)
                                {
                                    IsProcessSuccess = false;
                                    Logger.SaveLog("Modify Booking", DateTime.Now, DateTime.Now, "SegmentIdMapping is error", "");
                                    throw new ModifyBookingException("Add flight fail: Segment status invalid", "F011");
                                }
                                else
                                {
                                    // change flight success  if found seat cancel seat
                                    if (objSegmentIdMapping != null && objSegmentIdMapping.Count > 0)
                                    {
                                        booking.Mappings = CancelSeatInChangeFlight(objSegmentIdMapping, booking.Mappings);
                                    }
                                }
                            }
                            else
                            {
                                IsProcessSuccess = false;
                                // add flight fail

                                throw new ModifyBookingException("Add flight fail", "F011");
                            }
                        }//end add flight

                        //*********************************************************************

                        //Found SSR  set SSR to RS  Add Remark
                        if (IsProcessSuccess && services != null && services.Count > 0)
                        {
                            // if new seg  assign new seg id to ssr
                            if (objSegmentIdMapping != null && objSegmentIdMapping.Count > 0)
                            {
                                services.SetNewSegmentId(objSegmentIdMapping);
                            }
                            // convert ssr to RS
                            services.ToRecordset(ref rsService);

                            // add SSR remark 
                            IList<Remark> remarkList = AddRemark(strBookingId, agencyCode, strUserId);
                            remarkList.ToRecordset(ref rsRemark);

                        }// end ssr
                        //*********************************************************************
                        //Found SEAT and set seat to mapping
                        // return rs fee, mapping
                        if (IsProcessSuccess && seatAssigns != null && seatAssigns.Count > 0)
                        {
                            // set seat to mapping
                            booking.Mappings = SetSeatToMapping(booking.Mappings, seatAssigns, strUserId);

                          //  string xx = XMLHelper.Serialize(booking.Mappings, false);

                            //if change flight merge seat mapping to RS new mapping
                            foreach (Entity.Booking.Mapping m in booking.Mappings)
                            {
                                if (m.PassengerStatusRcd == "OK" && !string.IsNullOrEmpty(m.SeatNumber))
                                {
                                    if (objSegmentIdMapping != null && objSegmentIdMapping.Count > 0)
                                    {
                                        foreach (var element in objSegmentIdMapping)
                                        {
                                            // check seat to new mapping or not
                                            if (m.BookingSegmentId == new Guid(element.Key))
                                            {
                                                // assign seat to new mapping
                                                SetSeatToNewRsMapping(ref rsMapping, objSegmentIdMapping, m, strUserId);
                                            }
                                            else
                                            {
                                                // assign seat to old mapping
                                                SetSeatToSameRsMapping(ref rsMapping, m, strUserId);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // assign seat to old mapping
                                        SetSeatToSameRsMapping(ref rsMapping, m, strUserId);
                                    }
                                }
                            }

                            // ************ void seat fee if found double seat fee************
                            // void seat fee if new select case not paid
                            if (booking.Fees != null && booking.Fees.Count > 0)
                            {
                                VoidSeatFee(ref rsFees, seatAssigns, strUserId);
                            }


                        }// end seat
                        //*********************************************************************

                        //NameChange
                        if (IsProcessSuccess && NameChange != null && NameChange.Count > 0)
                        {
                            // updated name in mapping
                            SetNameToRsMapping(ref rsMapping, NameChange , strUserId);

                            //updated name in passenger
                            SetNameToRsPassenger(ref rsPassenger, NameChange, strUserId);
                        } 
                        // end name change
                        //*********************************************************************

                        //Add fee (bag + ssr + seat + name) to RS
                        if (IsProcessSuccess && fees != null && fees.Count > 0)
                        {

                            // if change flight set new seg id to fee
                            if (objSegmentIdMapping != null && objSegmentIdMapping.Count > 0)
                            {
                                fees = SetNewIdToFee(objSegmentIdMapping, fees);
                            }

                            // set user update
                            fees = AddCreateBy(fees,new Guid(strUserId));

                            // map fee to rs
                            fees.ToRecordset(ref rsFees);

                        }

                        // set pnrupd create by
                        // updated header new userid new time always
                        if (IsProcessSuccess)
                        {
                            UpdateFeeCreateBy(ref rsFees, strUserId);
                            UpdateMappingCreateBy(ref rsMapping, strUserId);
                            UpdateHeader(ref rsHeader, booking.Header, strUserId);
                            UpdateSegmentCreateBy(ref rsSegment, strUserId);
                            UpdateSSRCreateBy(ref rsService, strUserId);
                        }

                        //  when action code PRI : cal balance
                        // replace booking.fee from RSFee again
                        booking.Fees = booking.Fees.FillBooking(rsFees);
                       
                        booking.Mappings = booking.Mappings.FillBooking(rsMapping);

                        totalOutStanding = booking.CalOutStandingBalance(booking.Fees, booking.Mappings);


                        if (actionCode.Trim().ToUpper() == "PRI")
                        {
                            bookingResponse.Fees = bookingResponse.Fees.FillBooking(rsFees);
                            bookingResponse.Mappings = bookingResponse.Mappings.FillBooking(rsMapping);
                            bookingResponse.Services = bookingResponse.Services.FillBooking(rsService);
                        }
                        else if (actionCode.Trim().ToUpper() == "CON")
                        {
                            // save payment
                            if (IsProcessSuccess)
                            {
                                bool IsBookNow = false;

                                // for payment cal balance + allocate
                                booking = RsFillBooking(rsHeader, rsSegment, rsPassenger, rsFees, rsRemark, rsQuote, rsTax, rsService, rsMapping);

                                if (payments != null && payments.Count > 0)
                                {
                                    IsPaymentSuccess = COBPayment(booking, payments, strUserId, agencyCode, currencyRcd, totalOutStanding);
                                }
                                else
                                {
                                    IsBookNow = true;
                                }

                                // save booking
                                if (IsPaymentSuccess || IsBookNow)
                                {
                                    enumSaveError = objBooking.Save(ref rsHeader,
                                                                      ref rsSegment,
                                                                      ref rsPassenger,
                                                                      ref rsRemark,
                                                                      ref rsPayment,
                                                                      ref rsMapping,
                                                                      ref rsService,
                                                                      ref rsTax,
                                                                      ref rsFees,
                                                                      ref createTickets,
                                                                      ref readBooking,
                                                                      ref readOnly,
                                                                      ref bSetLock,
                                                                      ref bCheckSeatAssignment,
                                                                      ref bCheckSessionTimeOut);


                                    if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                                    {
                                        Boolean resultTicket = false;

                                        // for create ticket
                                        if (IsPaymentSuccess)
                                            resultTicket = objBooking.TicketCreate(agencyCode, strUserId, strBookingId);

                                        if (resultTicket || IsBookNow)
                                        {
                                            //read booking  
                                            if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                                                string.Empty,
                                                0,
                                                ref rsHeader,
                                                ref rsSegment,
                                                ref rsPassenger,
                                                ref rsRemark,
                                                ref rsPayment,
                                                ref rsMapping,
                                                ref rsService,
                                                ref rsTax,
                                                ref rsFees,
                                                ref rsQuote,
                                                false,
                                                false,
                                                "{" + strUserId.ToUpper() + "}",
                                                true,
                                                true,
                                                true,
                                                true,
                                                true,
                                                true,
                                                true,
                                                true,
                                                true,
                                                true,
                                                string.Empty,
                                                string.Empty) == true)
                                            {
                                                // convert rs to booking for preparing header and mapping  value
                                                if (rsHeader != null && rsHeader.RecordCount > 0)
                                                {
                                                    bookingResponse.Header = booking.Header.FillBooking(rsHeader);
                                                }
                                                if (rsSegment != null && rsSegment.RecordCount > 0)
                                                {
                                                    bookingResponse.Segments = booking.Segments.FillBooking(rsSegment);
                                                }
                                                if (rsPassenger != null && rsPassenger.RecordCount > 0)
                                                {
                                                    bookingResponse.Passengers = booking.Passengers.FillBooking(rsPassenger);
                                                }
                                                if (rsFees != null && rsFees.RecordCount > 0)
                                                {
                                                    bookingResponse.Fees = booking.Fees.FillBooking(rsFees);
                                                }
                                                if (rsRemark != null && rsRemark.RecordCount > 0)
                                                {
                                                    bookingResponse.Remarks = booking.Remarks.FillBooking(rsRemark);
                                                }
                                                if (rsQuote != null && rsQuote.RecordCount > 0)
                                                {
                                                    bookingResponse.Quotes = booking.Quotes.FillBooking(rsQuote);
                                                }
                                                if (rsTax != null && rsTax.RecordCount > 0)
                                                {
                                                    bookingResponse.Taxs = booking.Taxs.FillBooking(rsTax);
                                                }
                                                if (rsService != null && rsService.RecordCount > 0)
                                                {
                                                    bookingResponse.Services = booking.Services.FillBooking(rsService);
                                                }
                                                if (rsMapping != null && rsMapping.RecordCount > 0)
                                                {
                                                    bookingResponse.Mappings = booking.Mappings.FillBooking(rsMapping);
                                                }
                                                if (rsPayment != null && rsPayment.RecordCount > 0)
                                                {
                                                    bookingResponse.Payments = booking.Payments.FillBooking(rsPayment);
                                                }
                                            }
                                            else
                                            {
                                                strError = "Read booking fail";
                                                throw new ModifyBookingException("Read booking fail", "B008");
                                            }
                                        }
                                        else
                                        {
                                            strError = "CREATE TICKET FAIL";

                                            throw new ModifyBookingException("CREATE TICKET FAIL", "P006");
                                        }
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                                    {
                                        strError = "BOOKINGINUSE";
                                        throw new BookingSaveException("BOOKINGINUSE");
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                                    {
                                        strError = "BOOKINGREADERROR";
                                        throw new BookingSaveException("BOOKINGREADERROR");
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                                    {
                                        strError = "DATABASEACCESS";
                                        throw new BookingSaveException("DATABASEACCESS");
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                                    {
                                        strError = "DUPLICATESEAT";
                                        throw new BookingSaveException("Duplicated Seat", "277");

                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                                    {
                                        strError = "SESSIONTIMEOUT";
                                        throw new BookingSaveException("SESSIONTIMEOUT", "H005");
                                    }
                                } // end payment
                                else
                                {
                                    throw new ModifyBookingException("Payment fail", "P001");
                                }
                            }// end process
                            else
                            {
                                throw new ModifyBookingException("Modify booking fail", "300");
                            }
                        }
                        else
                        {
                            // invalid action code
                            strError = "Invalid ActionCode";
                            throw new ModifyBookingException("Invalid ActionCode", "");
                        }
                    }
                    else
                    {
                        throw new ModifyBookingException("Read booking fail", "B008");
                    }
                }
            }
            catch (ModifyBookingException)
            {
                throw;//new ModifyBookingException("Payment fail", "P001");
            }
            catch (BookingSaveException)
            {
                throw;//new ModifyBookingException("Payment fail", "P001");
            }
            catch
            {
                throw;// new ModifyBookingException(strError);
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);

            }
            return bookingResponse;
        }

        private bool readForSave(string strBookingId, string strUserId, 
                                    ref ADODB.Recordset rsHeader,
                                    ref ADODB.Recordset rsSegment, 
                                    ref ADODB.Recordset rsPassenger, 
                                    ref ADODB.Recordset rsRemark,
                                    ref ADODB.Recordset rsPayment, 
                                    ref  ADODB.Recordset rsMapping, 
                                    ref ADODB.Recordset rsService,
                                    ref  ADODB.Recordset rsTax, 
                                    ref  ADODB.Recordset rsFees, 
                                    ref  ADODB.Recordset rsQuote
            )
        {
            tikAeroProcess.Booking objBooking = null;
            bool success = false;

            if (_server.Length > 0)
            {
                Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                remote = null;
            }
            else
            {
                objBooking = new tikAeroProcess.Booking();
            }
            
            if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                                        string.Empty,
                                        0,
                                        ref rsHeader,
                                        ref rsSegment,
                                        ref rsPassenger,
                                        ref rsRemark,
                                        ref rsPayment,
                                        ref rsMapping,
                                        ref rsService,
                                        ref rsTax,
                                        ref rsFees,
                                        ref rsQuote,
                                        false,
                                        false,
                                        "{" + strUserId.ToUpper() + "}",
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        true,
                                        string.Empty,
                                        string.Empty) == true)
            {
                success = true;
            }

            return success;
        }

        private bool CancelSegment(string strBookingId, 
                                    string strUserId,
                                    string agencyCode,
                                    ref ADODB.Recordset rsSegment,
                                    ref ADODB.Recordset rsPayment,
                                    ref  ADODB.Recordset rsMapping,
                                    ref ADODB.Recordset rsService,
                                    ref  ADODB.Recordset rsTax,
                                    ref  ADODB.Recordset rsQuote,
                                    string bookingSegmentId

    )
        {
            tikAeroProcess.Booking objBooking = null;
            bool success = true;

            if (_server.Length > 0)
            {
                Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                remote = null;
            }
            else
            {
                objBooking = new tikAeroProcess.Booking();
            }

            bool processTicketRefund = true;
            bool IncludeRefundableOnly = true;
            DateTime dtNoShowSegmentDatetime = new DateTime();
            bool bReturnConfirmCancelStatus = true;
            bool bCancelSegment = false;
            bool bVoid = false;
            bool bWaveFee = false;

            // cancel segment
          //  foreach (FlightSegment f in segments)
           // {
              //  if (f.SegmentStatusRcd != null && f.SegmentStatusRcd.ToUpper() == "XX")
               // {
                    bCancelSegment = objBooking.SegmentCancel("{" + bookingSegmentId + "}",
                                                ref rsSegment,
                                                ref rsMapping,
                                                ref rsService,
                                                ref rsPayment,
                                                ref rsTax,
                                                ref rsQuote,
                                                ref strUserId,
                                                ref agencyCode,
                                                ref bWaveFee,
                                                ref processTicketRefund,
                                                ref IncludeRefundableOnly,
                                                ref dtNoShowSegmentDatetime,
                                                ref bVoid,
                                                ref bReturnConfirmCancelStatus);

                    if (bCancelSegment == false)
                    {
                        success = false;
                       // break;
                    }
             //   }
           // }
                            
            return success;
        }

        private IList<entity.Mapping> UpdatedSeatToMapping(IList<entity.Mapping> mappings, ref ADODB.Recordset rsMapping, string strUserId)
        {
            if (mappings == null )
                return new List<entity.Mapping>();

            List<entity.Mapping> rqstMappings = new List<entity.Mapping>();

            try
            {
                foreach (Mapping m in mappings)
                {
                    rsMapping.MoveFirst();

                    while (!rsMapping.EOF)
                    {

                        if (RecordsetHelper.ToGuid(rsMapping, "passenger_id") == m.PassengerId &&
                            RecordsetHelper.ToGuid(rsMapping, "booking_segment_id") == m.BookingSegmentId)
                        {
                            string currentSeat = rsMapping.Fields["seat_number"].Value as string;


                            if (!string.Equals(currentSeat, m.SeatNumber, StringComparison.OrdinalIgnoreCase))
                            {
                                rsMapping.Fields["seat_number"].Value = m.SeatNumber;

                                if (m.SeatNumber.Length > 1)
                                {

                                    rsMapping.Fields["seat_column"].Value = m.SeatNumber.Substring(m.SeatNumber.Length - 1, 1);
                                    rsMapping.Fields["seat_row"].Value = Convert.ToInt16(m.SeatNumber.Substring(0, m.SeatNumber.Length - 1));

                                   // m.SeatColumn = m.SeatNumber.Substring(m.SeatNumber.Length - 1, 1);
                                   // m.SeatRow = Convert.ToInt16(m.SeatNumber.Substring(0, m.SeatNumber.Length - 1));
                                }
                                else
                                {
                                    rsMapping.Fields["seat_column"].Value = DBNull.Value;
                                    rsMapping.Fields["seat_row"].Value = DBNull.Value;
                                 
                                }

                                rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                                rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                                rsMapping.Fields["create_date_time"].Value = DateTime.Now;
                                rsMapping.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                                m.FlightId = RecordsetHelper.ToGuid(rsMapping, "flight_id");
                                m.OriginRcd = RecordsetHelper.ToString(rsMapping, "origin_rcd");
                                m.DestinationRcd = RecordsetHelper.ToString(rsMapping, "destination_rcd");
                                m.BoardingClassRcd = RecordsetHelper.ToString(rsMapping, "boarding_class_rcd");

                                rqstMappings.Add(m);
                            }
                        }
                        rsMapping.MoveNext();
                    }
                }

            }
            catch (System.Exception ex)
            {

            }

            return rqstMappings;
        }

        private bool IsRemoveSeat(IList<entity.Mapping> rqstMappings)
        {
            bool isRemoveSeat = false;

            foreach (var m in rqstMappings)
            {
                if (string.IsNullOrEmpty(m.SeatNumber))
                {
                    isRemoveSeat = true;
                    break;
                }
            }
            return isRemoveSeat;
        }

        private bool IsRemoveSeatAll(IList<entity.Mapping> rqstMappings)
        {
            bool isRemoveSeatAll = false;
         
            if (rqstMappings.Count > 1)
            {
                int countSeatEmpty = 0;

                foreach (var m in rqstMappings)
                {
                    if (string.IsNullOrEmpty(m.SeatNumber))
                    {
                        countSeatEmpty = countSeatEmpty + 1;
                    }
                }

               if(countSeatEmpty > 1)
                {
                    isRemoveSeatAll = true;
                }
            }

            return isRemoveSeatAll;
        }

        private bool CaseNullPassengerChangeSeatMoreOne(IList<entity.Mapping> rqstMappings)
        {
            bool changeSeatNullPassenger = false;
           
            if (rqstMappings.Count > 1)
            {
                int countSeatEmpty = 0;
                foreach (var m in rqstMappings)
                {
                    if (!string.IsNullOrEmpty(m.SeatNumber))
                    {
                        countSeatEmpty = countSeatEmpty + 1;
                    }
                }

                if (countSeatEmpty > 1)
                {
                    changeSeatNullPassenger = true;
                }
            }

            return changeSeatNullPassenger;
        }


        private bool CaseNullPassengerOUT_Change(IList<entity.Mapping> rqstMappings)
        {
            if (rqstMappings.Count <= 1)
                return false;

            bool isNullPassengerOUT = false;
            bool isNullPassengerChange = false;
            int countSeatEmpty = 0, countSeatNotEmpty = 0;

            foreach (var m in rqstMappings)
            {
                if (string.IsNullOrEmpty(m.SeatNumber))
                {
                    countSeatEmpty++;
                    if (countSeatEmpty > 0)
                        isNullPassengerOUT = true;
                }
                else
                {
                    countSeatNotEmpty++;
                    if (countSeatNotEmpty > 0)
                        isNullPassengerChange = true;
                }

                // If both conditions are met, no need to continue looping
                if (isNullPassengerOUT && isNullPassengerChange)
                    return true;
            }

            return false;
        }

        private bool CaseNullPassengerChange_OUT(IList<entity.Mapping> rqstMappings)
        {
            if (rqstMappings.Count <= 1)
                return false;

            bool isNullPassengerFirstChange = false;
            bool isNullPassengerOUT = false;
            int countSeatEmpty = 0, countSeatNotEmpty = 0;


            if (!string.IsNullOrEmpty(rqstMappings[0].SeatNumber))
            {
                isNullPassengerFirstChange = true;
            }

            if (isNullPassengerFirstChange == true)
            {
                foreach (var m in rqstMappings)
                {
                    if (string.IsNullOrEmpty(m.SeatNumber))
                    {
                        countSeatNotEmpty++;
                        if (countSeatNotEmpty > 0)
                            isNullPassengerOUT = true;
                    }
                }
            }

                // If both conditions are met, no need to continue looping
                if (isNullPassengerFirstChange && isNullPassengerOUT)
                    return true;
            

            return false;
        }


        private bool IsFindRQST(IList<entity.Mapping> rqstMappings, IList<entity.Mapping> oldMapping, ref ADODB.Recordset rsService)
        {
            bool isFoundRQST = false;

            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_rcd = RecordsetHelper.ToString(rsService, "special_service_rcd");
                    string special_service_status_rcd = RecordsetHelper.ToString(rsService, "special_service_status_rcd");

                    foreach (var mapping in rqstMappings)
                    {
                        if (bookingSegmentId.Equals(mapping.BookingSegmentId)
                            && string.Equals(special_service_rcd, "RQST", StringComparison.OrdinalIgnoreCase)
                            && special_service_status_rcd == "HK"
                          )
                        {   // have RQST in system
                            isFoundRQST = true;
                            break;
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }


            //bool isNullPassengerAddAll = true;
            //foreach (var rqst in rqstMappings)
            //{
            //    if (string.IsNullOrEmpty(rqst.SeatNumber))
            //    {
            //        isNullPassengerAddAll = false;
            //        break;
            //    }
            //}

            //if(isFoundRQST == false && isNullPassengerAddAll == true)
            //{
            //    isAddNewRQST = true;
            //}

            return isFoundRQST;
        }

        private bool CaseNullPassengerOUT_Change2(IList<entity.Mapping> rqstMappings)
        {
            bool isNullPassengerOUT_Change = false;
            bool isNullPassengerOUT = false;
            bool isNullPassengerChange = false;

            // find seat OUT
            if (rqstMappings.Count > 1)
            {
                int countSeatEmpty = 0;
                foreach (var m in rqstMappings)
                {
                    if (string.IsNullOrEmpty(m.SeatNumber))
                    {
                        countSeatEmpty++;
                        if (countSeatEmpty > 0)
                        {
                            isNullPassengerOUT = true;
                            break;
                        }
                    }
                }
            }

            // find seat Change
            if (rqstMappings.Count > 1)
            {
                int countSeatNOTEmpty = 0;
                foreach (var m in rqstMappings)
                {
                    if (!string.IsNullOrEmpty(m.SeatNumber))
                    {
                        countSeatNOTEmpty++;
                        if (countSeatNOTEmpty > 0)
                        {
                            isNullPassengerChange = true;
                            break;
                        }
                    }
                }
            }

            // defind this both seat OUT and Change
            if(isNullPassengerOUT && isNullPassengerChange)
            {
                isNullPassengerOUT_Change = true;
            }

            return isNullPassengerOUT_Change;
        }

        private bool IsPassengerNull(IList<entity.Mapping> rqstMapping, ref ADODB.Recordset rsService)
        {
            bool isPassengerNull = true;

            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_rcd = RecordsetHelper.ToString(rsService, "special_service_rcd");
                    string special_service_status_rcd = RecordsetHelper.ToString(rsService, "special_service_status_rcd");

                    foreach (var mapping in rqstMapping)
                    {
                        if (bookingSegmentId.Equals(mapping.BookingSegmentId) 
                            && string.Equals(special_service_rcd, "RQST", StringComparison.OrdinalIgnoreCase) 
                            && special_service_status_rcd == "HK"
                          ) 
                        {//if p id have id  that passengerNull case is false
                            if (!passengerId.Equals(Guid.Empty))
                            {
                                isPassengerNull = false;
                                return isPassengerNull;
                            }
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }

            return isPassengerNull;
        }

        private bool IsPassengerNull2(IList<entity.Mapping> rqstMapping, ref ADODB.Recordset rsService)
        {
            bool isPassengerNull = false;

            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_rcd = RecordsetHelper.ToString(rsService, "special_service_rcd");
                    string special_service_status_rcd = RecordsetHelper.ToString(rsService, "special_service_status_rcd");

                    foreach (var mapping in rqstMapping)
                    {
                        if (bookingSegmentId.Equals(mapping.BookingSegmentId) && string.Equals(special_service_rcd, "RQST", StringComparison.OrdinalIgnoreCase)
                          )
                        {
                            //found p id == guid empty  Null = true
                            if (passengerId.Equals(Guid.Empty))
                            {
                                isPassengerNull = true;
                                return isPassengerNull;
                            }
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }

            return isPassengerNull;
        }


        private bool IsUseRQSTTogether(Guid RQST_booking_segment_id, ref ADODB.Recordset rsService)
        {
            bool isUseRQSTTogether = false;

            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_rcd = RecordsetHelper.ToString(rsService, "special_service_rcd");
                    string special_service_status_rcd = RecordsetHelper.ToString(rsService, "special_service_status_rcd");


                        if (bookingSegmentId.Equals(RQST_booking_segment_id) 
                        && string.Equals(special_service_rcd, "RQST", StringComparison.OrdinalIgnoreCase)
                        && special_service_status_rcd == "HK"
                          )
                        {
                            //found p id == guid empty  Null = true
                            if (passengerId.Equals(Guid.Empty))
                            {
                            isUseRQSTTogether = true;
                                return isUseRQSTTogether;
                            }
                        }


                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }

            return isUseRQSTTogether;
        }

        private List<SegmentSeatInfo> AddSeatCharacteristicForPassNullWithRestPassenger(IList<entity.Mapping> rqstMapping, IList<entity.Mapping> oldMapping)
        {
            List<SegmentSeatInfo> resultList = new List<SegmentSeatInfo>();
            Dictionary<Guid, StringBuilder> segmentSeatMap = new Dictionary<Guid, StringBuilder>();

            try
            {
                foreach (var m in oldMapping)
                {
                    Guid bookingSegmentId = m.BookingSegmentId;
                    Guid passengerId = m.PassengerId;
                    int SeatRow = m.SeatRow;
                    string SeatColumn = m.SeatColumn;
                    string SeatNumber = m.SeatNumber;

                    foreach (var rqstM in rqstMapping)
                    {
                        if (bookingSegmentId.Equals(rqstM.BookingSegmentId)
                            && !string.IsNullOrEmpty(rqstM.SeatNumber)
                            && passengerId != rqstM.PassengerId
                            )
                        {
                            string seatCharacteristic = GetSeatCharacteristic(
                                rqstM.FlightId.ToString(),
                                rqstM.OriginRcd,
                                rqstM.DestinationRcd,
                                rqstM.BoardingClassRcd,
                                SeatRow,
                                SeatColumn
                            ) ?? string.Empty;

                            // Store seat info grouped by segment ID
                            if (!segmentSeatMap.ContainsKey(bookingSegmentId))
                            {
                                segmentSeatMap[bookingSegmentId] = new StringBuilder();
                            }

                            segmentSeatMap[bookingSegmentId].Append(SeatNumber).Append(seatCharacteristic);
                        }
                    }

                }

                // Convert dictionary to list of SegmentSeatInfo objects
                foreach (var entry in segmentSeatMap)
                {
                    resultList.Add(new SegmentSeatInfo
                    {
                        BookingSegmentId = entry.Key,
                        SeatDetails = entry.Value.ToString()
                    });
                }
            }
            catch (System.Exception ex)
            {

            }

            return resultList;
        }
        private List<SegmentSeatInfo> AddSeatCharacteristicForPassNullWithRestPassenger2(IList<entity.Mapping> rqstMapping)
        {
            List<SegmentSeatInfo> resultList = new List<SegmentSeatInfo>();
            Dictionary<Guid, StringBuilder> segmentSeatMap = new Dictionary<Guid, StringBuilder>();

            try
            {

                    foreach (var rqstM in rqstMapping)
                    {
                        if (!string.IsNullOrEmpty(rqstM.SeatNumber) )
                        {
                            string seatCharacteristic = GetSeatCharacteristic(
                                rqstM.FlightId.ToString(),
                                rqstM.OriginRcd,
                                rqstM.DestinationRcd,
                                rqstM.BoardingClassRcd,
                                rqstM.SeatRow,
                                rqstM.SeatColumn
                            ) ?? string.Empty;

                            // Store seat info grouped by segment ID
                            if (!segmentSeatMap.ContainsKey(rqstM.BookingSegmentId))
                            {
                                segmentSeatMap[rqstM.BookingSegmentId] = new StringBuilder();
                            }

                            segmentSeatMap[rqstM.BookingSegmentId].Append(string.Concat(rqstM.SeatRow, rqstM.SeatColumn)).Append(seatCharacteristic);
                        }
                    }

                // Convert dictionary to list of SegmentSeatInfo objects
                foreach (var entry in segmentSeatMap)
                {
                    resultList.Add(new SegmentSeatInfo
                    {
                        BookingSegmentId = entry.Key,
                        SeatDetails = entry.Value.ToString()
                    });
                }
            }
            catch (System.Exception ex)
            {

            }

            return resultList;
        }

        private string GetSeatCharacteristic(string flightId, string originRcd, string destinationRcd, string boardingClassRcd, int seatRow, string seatColumn)
        {
            string connectionString = ConfigHelper.ToString("SQLConnectionString");
            string seatCharacteristic = string.Empty;

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    string query = @"SELECT dbo.GetSeatCharacteristic(@FlightId, @OriginRcd, @DestinationRcd, @BoardingClassRcd, @SeatRow, @SeatColumn)";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@FlightId", flightId);
                        cmd.Parameters.AddWithValue("@OriginRcd", originRcd);
                        cmd.Parameters.AddWithValue("@DestinationRcd", destinationRcd);
                        cmd.Parameters.AddWithValue("@BoardingClassRcd", boardingClassRcd);
                        cmd.Parameters.AddWithValue("@SeatRow", seatRow);
                        cmd.Parameters.AddWithValue("@SeatColumn", seatColumn);

                        conn.Open();
                        object result = cmd.ExecuteScalar();

                        if (result != null && result != DBNull.Value)
                        {
                            seatCharacteristic = result.ToString();
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {

            }
            catch (System.Exception ex)
            {

            }

            return seatCharacteristic;
        }

        private void  SetSeatCharacteristic(ref IList<entity.PassengerService> listOfRqst , List<SegmentSeatInfo> listOfSeatCharacteristic)
        {
            if (listOfRqst != null && listOfRqst.Count > 0)
            {
                foreach (var seat in listOfSeatCharacteristic)
                {
                    foreach (var ps in listOfRqst)
                    {
                        if (ps.BookingSegmentId == seat.BookingSegmentId)
                        {
                            ps.ServiceText += seat.SeatDetails;
                        }

                    }

                }
            }

        }




        private IList<entity.PassengerService> CreateSeatRQST(IList<entity.Mapping> rqstMapping, short numberOfPassenger )
        {
            List<entity.PassengerService> services = new List<entity.PassengerService>();

            foreach (var mapping in rqstMapping)
            {
                if (!string.IsNullOrEmpty(mapping.SeatNumber)) // found seat so create RQST
                {
                    // call GetSeatCharacteristic
                    string seatCharacteristic = GetSeatCharacteristic(mapping.FlightId.ToString(), mapping.OriginRcd, mapping.DestinationRcd,
                        mapping.BoardingClassRcd, mapping.SeatRow, mapping.SeatColumn);

                    entity.PassengerService service = new entity.PassengerService
                    {
                        PassengerSegmentServiceId = Guid.NewGuid(),
                        PassengerId = mapping.PassengerId,
                        BookingSegmentId = mapping.BookingSegmentId,
                        NumberOfUnits = 1,
                        ServiceOnRequestFlag = 0,
                        SpecialServiceRcd = "RQST",
                        ServiceText = string.Concat(mapping.SeatNumber, seatCharacteristic),
                        SpecialServiceStatusRcd = "HK",
                        SpecialServiceChangeStatusRcd = "KK",
                        UnitRcd = string.Empty,

                        OriginRcd = mapping.OriginRcd,
                        DestinationRcd = mapping.DestinationRcd,

                        CreateDateTime = DateTime.Now,
                        CreateBy = mapping.UpdateBy,

                        UpdateDateTime = DateTime.Now,
                        UpdateBy = mapping.UpdateBy
                    };

                    services.Add(service);
                }
            }
            
            return services;
        }

        private IList<entity.PassengerService> CreateSeatRQSTNullPassenger_OUT_Change(IList<entity.Mapping> rqstMapping, short numberOfPassenger, ref ADODB.Recordset rsService)
        {
            // need to delete  the RQST the active old one first

            //Guid p_id = rqstMapping[0].PassengerId;

            //entity.Booking b = new Booking();
            //List<entity.PassengerService> servicesTmp = new List<PassengerService>();
            //b.Services = servicesTmp.FillBooking(rsService);

            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_status_rcd = rsService.Fields["special_service_status_rcd"].Value?.ToString();
                    string special_service_rcd = rsService.Fields["special_service_rcd"].Value?.ToString();

                    foreach (var m in rqstMapping)
                    {
                        if (
                            bookingSegmentId == m.BookingSegmentId &&
                            (passengerId == m.PassengerId) &&
                           // string.Equals(special_service_status_rcd, "HK", StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(special_service_rcd, "RQST", StringComparison.OrdinalIgnoreCase)
                            )
                        {
                            rsService.Delete();
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {
                //
            }

            //entity.Booking b2 = new Booking();
            //List<entity.PassengerService> servicesTmp2 = new List<PassengerService>();
            //b2.Services = servicesTmp2.FillBooking(rsService);


            List<entity.PassengerService> services = new List<entity.PassengerService>();

            foreach (var mapping in rqstMapping)
            {
                if (!string.IsNullOrEmpty(mapping.SeatNumber)) // found seat so create RQST
                {
                    // call GetSeatCharacteristic
                    string seatCharacteristic = GetSeatCharacteristic(mapping.FlightId.ToString(), mapping.OriginRcd, mapping.DestinationRcd,
                        mapping.BoardingClassRcd, mapping.SeatRow, mapping.SeatColumn);

                    entity.PassengerService service = new entity.PassengerService
                    {
                        PassengerSegmentServiceId = Guid.NewGuid(),
                        PassengerId = mapping.PassengerId,
                        BookingSegmentId = mapping.BookingSegmentId,
                        NumberOfUnits = 1,
                        ServiceOnRequestFlag = 0,
                        SpecialServiceRcd = "RQST",
                        ServiceText = string.Concat(mapping.SeatNumber, seatCharacteristic),
                        SpecialServiceStatusRcd = "HK",
                        SpecialServiceChangeStatusRcd = "KK",
                        UnitRcd = string.Empty,

                        OriginRcd = mapping.OriginRcd,
                        DestinationRcd = mapping.DestinationRcd,

                        CreateDateTime = DateTime.Now,
                        CreateBy = mapping.UpdateBy,

                        UpdateDateTime = DateTime.Now,
                        UpdateBy = mapping.UpdateBy
                    };

                    services.Add(service);
                }
            }

            return services;
        }

        private void CancelSeatRQSTForRemoveSeatAllOfPassengerNull(IList<entity.Mapping> rqstMappings, ref ADODB.Recordset rsService, string strUserId)
        {
            try
            {

                // Clone the recordset
                ADODB.Recordset rsServiceTmp = rsService.Clone();
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    foreach (var rqstMapping in rqstMappings)
                    {
                        if (IsMatchingSpecialService(rsService, rqstMapping))
                        {
                            rsServiceTmp.MoveFirst();
                            while (!rsServiceTmp.EOF)
                            {
                                if (IsMatchingSpecialServiceForDelete(rsService, rsServiceTmp))
                                {
                                    rsService.Delete();
                                    break;
                                }
                                else
                                {
                                    UpdateSpecialService(rsService, strUserId);
                                }

                                rsServiceTmp.MoveNext();
                            }

                            break;
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {
            }
        }

        private bool IsMatchingSpecialService(ADODB.Recordset rsService, entity.Mapping rqstMapping)
        {
            return RecordsetHelper.ToGuid(rsService, "booking_segment_id") == rqstMapping.BookingSegmentId &&
                   RecordsetHelper.ToGuid(rsService, "passenger_id") == rqstMapping.PassengerId &&
                   string.Equals(rsService.Fields["special_service_status_rcd"].Value?.ToString(), "HK", StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(rsService.Fields["special_service_rcd"].Value?.ToString(), "RQST", StringComparison.OrdinalIgnoreCase);
        }

        private bool IsMatchingSpecialServiceForDelete(ADODB.Recordset rsService, ADODB.Recordset rsServiceTmp)
        {
            return RecordsetHelper.ToGuid(rsServiceTmp, "booking_segment_id") == RecordsetHelper.ToGuid(rsService, "booking_segment_id") &&
                   RecordsetHelper.ToGuid(rsServiceTmp, "passenger_id") == Guid.Empty &&
                   string.Equals(rsServiceTmp.Fields["special_service_status_rcd"].Value?.ToString(), "XX", StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(rsServiceTmp.Fields["special_service_rcd"].Value?.ToString(), "RQST", StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(rsService.Fields["special_service_change_status_rcd"].Value?.ToString(), "UC", StringComparison.OrdinalIgnoreCase);
        }

        private void UpdateSpecialService(ADODB.Recordset rsService, string strUserId)
        {
            rsService.Fields["update_date_time"].Value = (object)DateTime.Now ?? DBNull.Value;
            rsService.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
            rsService.Fields["special_service_status_rcd"].Value = "XX";
            rsService.Fields["special_service_change_status_rcd"].Value = "UC";
        }

        private void CancelSeatRQSTNullPassenger(IList<entity.Mapping> rqstMappings, ref ADODB.Recordset rsService, string strUserId)
        {
            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_status_rcd = rsService.Fields["special_service_status_rcd"].Value?.ToString();
                    string special_service_rcd = rsService.Fields["special_service_rcd"].Value?.ToString();

                    foreach (var rqstMapping in rqstMappings)
                    {
                        if (
                            bookingSegmentId == rqstMapping.BookingSegmentId &&
                            passengerId == Guid.Empty &&
                            string.Equals(special_service_status_rcd, "HK", StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(special_service_rcd, "RQST", StringComparison.OrdinalIgnoreCase)
                            )
                        {
                            rsService.Fields["update_date_time"].Value = DateTime.Now;
                            rsService.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                            rsService.Fields["special_service_status_rcd"].Value = "XX";
                            rsService.Fields["special_service_change_status_rcd"].Value = "UC";
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {
                //
            }
        }

        private void CancelSeatRQSTNormal(IList<entity.Mapping> rqstMappings, ref ADODB.Recordset rsService, string strUserId)
        {
            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_status_rcd = rsService.Fields["special_service_status_rcd"].Value?.ToString();
                    string special_service_rcd = rsService.Fields["special_service_rcd"].Value?.ToString();

                    foreach (var rqstMapping in rqstMappings)
                    {
                        if (
                            bookingSegmentId == rqstMapping.BookingSegmentId &&
                            (passengerId == rqstMapping.PassengerId) &&
                            string.Equals(special_service_status_rcd, "HK", StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(special_service_rcd, "RQST", StringComparison.OrdinalIgnoreCase)
                            )
                        {
                            rsService.Fields["update_date_time"].Value = DateTime.Now;
                            rsService.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                            rsService.Fields["special_service_status_rcd"].Value = "XX";
                            rsService.Fields["special_service_change_status_rcd"].Value = "UC"; // for ASVC 
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {
                //
            }
        }

       
        private void UpdateStatusKDtoASVC(IList<entity.Mapping> rqstMapping, ref ADODB.Recordset rsService)
        {
            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_rcd = RecordsetHelper.ToString(rsService, "special_service_rcd");

                    foreach (var mapping in rqstMapping)
                    {
                        if (bookingSegmentId.Equals(mapping.BookingSegmentId) && string.Equals(special_service_rcd, "ASVC", StringComparison.OrdinalIgnoreCase) )
                        {
                            if (passengerId.Equals(mapping.PassengerId))
                            {
                                rsService.Fields["special_service_change_status_rcd"].Value = "KD";
                            }
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        private void CancelASVC(IList<entity.Mapping> rqstMapping, ref ADODB.Recordset rsService)
        {
            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_rcd = RecordsetHelper.ToString(rsService, "special_service_rcd");

                    foreach (var mapping in rqstMapping)
                    {
                        if (bookingSegmentId.Equals(mapping.BookingSegmentId) && string.Equals(special_service_rcd, "ASVC", StringComparison.OrdinalIgnoreCase))
                        {
                            if (passengerId.Equals(mapping.PassengerId))
                            {
                                rsService.Fields["special_service_change_status_rcd"].Value = "KD";
                                rsService.Fields["special_service_status_rcd"].Value = "XX";
                            }
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }
        }


        private IList<entity.PassengerService> CreateSeatRQSTForPassengerNull(IList<entity.Mapping> rqstMapping, ref ADODB.Recordset rsService, short numberOfPassenger,bool isChangeSeatNullPassengerMoreOne, bool isChangeSeatNullPassengerSwitchSeatLastPax)
        {
            List<entity.PassengerService> services = new List<entity.PassengerService>();

            //entity.Booking b = new Booking();
            //List<entity.PassengerService> servicesTmp = new List<PassengerService>();

            //b.Services = servicesTmp.FillBooking(rsService);
            //entity.PassengerService latestService = null;
            //Guid lastPassengerSegmentServiceId = Guid.Empty;

            //if (b.Services != null && b.Services.Count > 0)
            //{
            //    latestService = b.Services[b.Services.Count - 1];
            //    lastPassengerSegmentServiceId = latestService.PassengerSegmentServiceId;
            //}

            //if (isChangeSeatNullPassengerSwitchSeatLastPax)
            //{

            //        rsService.MoveFirst();

            //        while (!rsService.EOF)
            //        {
            //            Guid passengerSegmentServiceId = RecordsetHelper.ToGuid(rsService, "passenger_segment_service_id");

            //            if (passengerSegmentServiceId == lastPassengerSegmentServiceId)
            //            {
            //                rsService.Delete();
            //                break; 
            //            }

            //            rsService.MoveNext();
            //        }

            //}

            // for test
            //entity.Booking bb = new Booking();
            //List<entity.PassengerService> servicesTmpp = new List<PassengerService>();

            //bb.Services = servicesTmpp.FillBooking(rsService);


            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    short numberOfUnits = RecordsetHelper.ToInt16(rsService, "number_of_units");
                    Guid passenger_id = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string specialServiceRcd = RecordsetHelper.ToString(rsService, "special_service_rcd");
                    string specialServiceStatusRcd = RecordsetHelper.ToString(rsService, "special_service_status_rcd");
                    string specialServiceChangeStatusRcd = RecordsetHelper.ToString(rsService, "special_service_change_status_rcd");
                    Guid id = RecordsetHelper.ToGuid(rsService, "passenger_segment_segment_id");

                    foreach (var mapping in rqstMapping)
                    {
                        if (bookingSegmentId.Equals(mapping.BookingSegmentId) && !string.IsNullOrEmpty(mapping.SeatNumber))
                        {

                            if (isChangeSeatNullPassengerMoreOne)
                            {
                                if(isChangeSeatNullPassengerSwitchSeatLastPax)
                                {
                                    entity.PassengerService service = new entity.PassengerService
                                    {
                                        PassengerSegmentServiceId = Guid.NewGuid(),
                                        BookingSegmentId = bookingSegmentId,
                                        PassengerId = Guid.Empty,
                                        NumberOfUnits = (short)mapping.NumberOfUnits,
                                        ServiceOnRequestFlag = 0,
                                        SpecialServiceRcd = "RQST",
                                        ServiceText = string.Empty,
                                        SpecialServiceStatusRcd = "HK",
                                        SpecialServiceChangeStatusRcd = "KK",
                                        UnitRcd = string.Empty,
                                        OriginRcd = mapping.OriginRcd,
                                        DestinationRcd = mapping.DestinationRcd,
                                        CreateDateTime = DateTime.Now,
                                        CreateBy = mapping.UpdateBy,
                                        UpdateDateTime = DateTime.Now,
                                        UpdateBy = mapping.UpdateBy
                                    };
                                    services.Add(service);
                                }
                            }                               
                            else
                            {
                                // call GetSeatCharacteristic
                                string seatCharacteristic = GetSeatCharacteristic(mapping.FlightId.ToString(), mapping.OriginRcd, mapping.DestinationRcd,
                                    mapping.BoardingClassRcd, mapping.SeatRow, mapping.SeatColumn);

                                entity.PassengerService service = new entity.PassengerService
                                {
                                    PassengerSegmentServiceId = Guid.NewGuid(),
                                    BookingSegmentId = bookingSegmentId,
                                    PassengerId = Guid.Empty,
                                    NumberOfUnits = (short)mapping.NumberOfUnits,
                                    ServiceOnRequestFlag = 0,
                                    SpecialServiceRcd = "RQST",
                                    ServiceText = string.Concat(mapping.SeatNumber, seatCharacteristic),
                                    SpecialServiceStatusRcd = "HK",
                                    SpecialServiceChangeStatusRcd = "KK",
                                    UnitRcd = string.Empty,
                                    OriginRcd = mapping.OriginRcd,
                                    DestinationRcd = mapping.DestinationRcd,
                                    CreateDateTime = DateTime.Now,
                                    CreateBy = mapping.UpdateBy,
                                    UpdateDateTime = DateTime.Now,
                                    UpdateBy = mapping.UpdateBy
                                };
                                services.Add(service);
                            }
                               

                            return services;
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }

            return services;
        }


        private void UpdateStatusKDASCVPassengerNull(IList<entity.Mapping> rqstMapping, ref ADODB.Recordset rsService)
        {
            try
            {
                rsService.MoveFirst();

                while (!rsService.EOF)
                {
                    Guid bookingSegmentId = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_rcd = RecordsetHelper.ToString(rsService, "special_service_rcd");

                    foreach (var mapping in rqstMapping)
                    {
                        if (bookingSegmentId.Equals(mapping.BookingSegmentId) //match all passengers
                            && string.Equals(special_service_rcd, "ASVC", StringComparison.OrdinalIgnoreCase)
                            
                            )
                        {
                            rsService.Fields["special_service_change_status_rcd"].Value = "KD";
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }

        }


        private IList<entity.PassengerService> CreateRQSTSeparatePassengers(IList<entity.Mapping> rqstMapping, IList<entity.Mapping> oldMappings)
        {
            List<entity.PassengerService> services = new List<entity.PassengerService>();

            foreach (var oldMapping in oldMappings)
            {
                Guid bookingSegmentId = oldMapping.BookingSegmentId;
                Guid passengerId = oldMapping.PassengerId;

                foreach (var rqstM in rqstMapping)
                {
                    if (bookingSegmentId.Equals(rqstM.BookingSegmentId)
                        && passengerId != rqstM.PassengerId
                        )
                    {
                        // call GetSeatCharacteristic
                        string seatCharacteristic = GetSeatCharacteristic(oldMapping.FlightId.ToString(), oldMapping.OriginRcd, oldMapping.DestinationRcd,
                            oldMapping.BoardingClassRcd, oldMapping.SeatRow, oldMapping.SeatColumn);

                        entity.PassengerService service = new entity.PassengerService
                        {
                            PassengerSegmentServiceId = Guid.NewGuid(),
                            BookingSegmentId = bookingSegmentId,
                            PassengerId = oldMapping.PassengerId,
                            NumberOfUnits = 1,
                            ServiceOnRequestFlag = 0,
                            SpecialServiceRcd = "RQST",
                            ServiceText = string.Concat(oldMapping.SeatNumber, seatCharacteristic),
                            SpecialServiceStatusRcd = "HK",
                            SpecialServiceChangeStatusRcd = "KK",
                            UnitRcd = string.Empty,
                            OriginRcd = oldMapping.OriginRcd,
                            DestinationRcd = oldMapping.DestinationRcd,
                            CreateDateTime = DateTime.Now,
                            CreateBy = oldMapping.UpdateBy,
                            UpdateDateTime = DateTime.Now,
                            UpdateBy = oldMapping.UpdateBy
                        };

                        services.Add(service);
                    }
                }
            }

            return services;
        }

        private void UpdateStatusKDRSeparatePassengers(IList<entity.Mapping> rqstMapping, IList<entity.Mapping> oldMappings, ref ADODB.Recordset rsService, string strUserId)
        {
            rsService.MoveFirst();

            try
            {
                while (!rsService.EOF)
                {
                    Guid bookingSegmentIdRS = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerIdRS = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_status_rcd = rsService.Fields["special_service_status_rcd"].Value?.ToString();
                    string special_service_rcd = rsService.Fields["special_service_rcd"].Value?.ToString();

                    foreach (var oldMapping in oldMappings)
                    {
                        Guid bookingSegmentIdM = oldMapping.BookingSegmentId;
                        Guid passengerIdM = oldMapping.PassengerId;

                        foreach (var rqstM in rqstMapping)
                        {
                            if (bookingSegmentIdM.Equals(rqstM.BookingSegmentId)
                                && !passengerIdM.Equals(rqstM.PassengerId)
                                && bookingSegmentIdRS.Equals(bookingSegmentIdM)
                                && passengerIdRS.Equals(passengerIdM)
                                && special_service_rcd.Equals("ASVC")
                                )
                            {
                                rsService.Fields["special_service_change_status_rcd"].Value = "KD";

                                break;

                            }
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        private void CancelRQSTSeparatePassengers(IList<entity.Mapping> rqstMapping, IList<entity.Mapping> oldMappings, ref ADODB.Recordset rsService, string strUserId)
        {
            rsService.MoveFirst();

            try
            {
                while (!rsService.EOF)
                {
                    Guid bookingSegmentIdRS = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerIdRS = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_status_rcd = rsService.Fields["special_service_status_rcd"].Value?.ToString();
                    string special_service_rcd = rsService.Fields["special_service_rcd"].Value?.ToString();

                    foreach (var oldMapping in oldMappings)
                    {
                        Guid bookingSegmentIdM = oldMapping.BookingSegmentId;
                        Guid passengerIdM = oldMapping.PassengerId;

                        foreach (var rqstM in rqstMapping)
                        {
                            if (bookingSegmentIdM.Equals(rqstM.BookingSegmentId)
                                // && passengerIdM.Equals(rqstM.PassengerId)
                              && passengerIdRS == Guid.Empty
                                && bookingSegmentIdRS.Equals(bookingSegmentIdM)
                                && special_service_status_rcd.Equals("HK")
                                && special_service_rcd.Equals("RQST")
                                )
                            {
                                rsService.Fields["update_date_time"].Value = DateTime.Now;
                                rsService.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                                rsService.Fields["special_service_status_rcd"].Value = "XX";
                                rsService.Fields["special_service_change_status_rcd"].Value = "UC";

                                break;

                            }
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }         
        }

        private void CancelASVCSeparatePassengers(IList<entity.Mapping> rqstMapping, IList<entity.Mapping> oldMappings, ref ADODB.Recordset rsService, string strUserId)
        {
            rsService.MoveFirst();

            try
            {
                while (!rsService.EOF)
                {
                    Guid bookingSegmentIdRS = RecordsetHelper.ToGuid(rsService, "booking_segment_id");
                    Guid passengerIdRS = RecordsetHelper.ToGuid(rsService, "passenger_id");
                    string special_service_status_rcd = rsService.Fields["special_service_status_rcd"].Value?.ToString();
                    string special_service_rcd = rsService.Fields["special_service_rcd"].Value?.ToString();

                    foreach (var oldMapping in oldMappings)
                    {
                        Guid bookingSegmentIdM = oldMapping.BookingSegmentId;
                        Guid passengerIdM = oldMapping.PassengerId;

                        foreach (var rqstM in rqstMapping)
                        {
                            if (bookingSegmentIdM.Equals(rqstM.BookingSegmentId)
                                 && passengerIdM.Equals(rqstM.PassengerId)
                                 && passengerIdRS.Equals(rqstM.PassengerId)
                                 && bookingSegmentIdRS.Equals(bookingSegmentIdM)
                                 && special_service_rcd.Equals("ASVC")
                                 )
                            {
                                rsService.Fields["special_service_change_status_rcd"].Value = "KD";
                                rsService.Fields["special_service_status_rcd"].Value = "XX";
                                break;

                            }
                        }
                    }

                    rsService.MoveNext();
                }
            }
            catch (System.Exception ex)
            {

            }
        }


        private IList<entity.PassengerService> moveSSR(Booking booking, List<KeyValuePair<string, string>> objSegmentIdMapping, ref  ADODB.Recordset rsService, IList<entity.FlightSegment> segments, IList<entity.PassengerService> services)
        {
            if (services == null)
                services = new List<entity.PassengerService>();

            foreach (FlightSegment oldF in segments)
            {
                foreach (FlightSegment newF in segments)
                {
                    foreach (var element in objSegmentIdMapping)
                    {

                        if (oldF.SegmentStatusRcd.ToUpper() == "XX" && newF.SegmentStatusRcd.ToUpper() == "NN" && oldF.BookingSegmentId == new Guid(element.Key) && newF.BookingSegmentId == new Guid(element.Value) && oldF.OriginRcd == newF.OriginRcd && oldF.DestinationRcd == newF.DestinationRcd)
                        {
                            // action move SSR
                            rsService.MoveFirst();
                            while (!rsService.EOF)
                            {
                                if (RecordsetHelper.ToGuid(rsService, "booking_segment_id") == oldF.BookingSegmentId)
                                {
                                    PassengerService ps = new PassengerService();
                                    ps.BookingSegmentId = newF.BookingSegmentId;
                                    ps.PassengerId = RecordsetHelper.ToGuid(rsService, "passenger_id");
                                    ps.ServiceText = RecordsetHelper.ToString(rsService, "service_text");
                                    ps.SpecialServiceRcd = RecordsetHelper.ToString(rsService, "special_service_rcd");
                                    ps.SpecialServiceStatusRcd = RecordsetHelper.ToString(rsService, "special_service_status_rcd");
                                    ps.NumberOfUnits = RecordsetHelper.ToInt16(rsService, "number_of_units");
                                    ps.PassengerSegmentServiceId = Guid.NewGuid();

                                    services.Add(ps);
                                }
                                rsService.MoveNext();
                            }
                        }
                    }
                }
            }

            return services;
        }

        private Booking AddNewMappingToBooking(Booking booking, IList<entity.Mapping> mappings, IList<entity.FlightSegment> segments, string agencyCode)
        {
            // add New Mapping to booking
            if (booking.Mappings == null)
                booking.Mappings = new List<Mapping>();

            if (mappings != null && mappings.Count > 0)
            {
                foreach (FlightSegment f in segments)
                {
                    foreach (Mapping m in mappings)
                    {
                        foreach (Passenger p in booking.Passengers)
                        {
                            if (f.BookingSegmentId == m.BookingSegmentId && m.PassengerId == p.PassengerId && f.SegmentStatusRcd.ToUpper() == "NN")
                            {
                                m.Firstname = p.Firstname;
                                m.Lastname = p.Lastname;
                                m.PassengerTypeRcd = p.PassengerTypeRcd;
                                m.PassengerStatusRcd = "OK";

                                m.AirlineRcd = f.AirlineRcd;
                                m.FlightNumber = f.FlightNumber;
                                m.FlightId = f.FlightId;
                                m.FlightConnectionId = f.FlightConnectionId;
                                m.DepartureDate = f.DepartureDate;
                                m.DepartureTime = f.DepartureTime;
                                m.ArrivalDate = f.ArrivalDate;
                                m.ArrivalTime = f.ArrivalTime;
                                m.JourneyTime = f.JourneyTime;
                                m.OriginRcd = f.OriginRcd;
                                m.DestinationRcd = f.DestinationRcd;
                                m.OdOriginRcd = f.OdOriginRcd;
                                m.OdDestinationRcd = f.OdDestinationRcd;
                                m.AgencyCode = agencyCode;

                                booking.Mappings.Add(m);
                            }
                        }
                    }
                }
            }

            return booking;
        }


        public entity.Booking OrderingBooking(string strBookingId,
                                    IList<entity.FlightSegment> segments,
                                    IList<entity.Mapping> mappings,
                                    IList<entity.Fee> fees,
                                    IList<entity.PassengerService> services,
                                    IList<entity.Tax> taxes,
                                    string strUserId,
                                    string agencyCode,
                                    string currencyRcd,
                                    string languageRcd)
        {
            tikAeroProcess.Booking objBooking = null;
            tikAeroProcess.Payment objPayment = null;
            tikAeroProcess.Fees objFees = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            ADODB.Recordset rsFlights = RecordsetHelper.FabricateFlightRecordset();

            string strXml = string.Empty;
            bool createTickets = false;
            bool readBooking = false;
            bool readOnly = false;
            bool bSetLock = false;
            // can not check dup seat in save booking because we save payment already
            bool bCheckSeatAssignment = true;
            bool bCheckSessionTimeOut = false;
            string paymentType = string.Empty;
            entity.Booking bookingResponse = new entity.Booking();
            entity.Booking booking = new entity.Booking();
            var objSegmentIdMapping = new List<KeyValuePair<string, string>>();
            bool IsProcessSuccess = true;
           // decimal totalOutStanding = 0;
            string strError = string.Empty;
            tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;
            bool addSegmentFlag = false;
            //bool changeFlightFlag = false;
            bool changedSegmentFlag = false;
            bool readSuccess = false;
            ArrayList alNewSegment = new ArrayList();
            var listSegmentIdMappingChangeFlight = new List<KeyValuePair<string, string>>();
            var SegmentIdMappingChangeFlight = new List<KeyValuePair<string, string>>();
            bool foundInfant = false;
            bool IsTTYMessage = false;
            bool IsRQSTSeat = false;

            short numberOfPassenger = 0;

            //ArrayList listSegmentIdMappingChangeFlight = new ArrayList();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                    //remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    //objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    //remote = null;
                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                    //objFees = new tikAeroProcess.Fees();
                    //objPayment = new tikAeroProcess.Payment();
                }

                if (!String.IsNullOrEmpty(strBookingId))
                {
                    //read booking first for prepare RS 
                    readSuccess = readForSave(strBookingId, strUserId, ref rsHeader, ref rsSegment, ref rsPassenger, ref rsRemark, ref rsPayment, ref rsMapping, ref rsService, ref rsTax, ref rsFees, ref rsQuote);

                    if (readSuccess)
                    {
                        // convert rs to booking for preparing header and mapping  value
                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);

                            if(booking.Header != null && !string.IsNullOrEmpty(booking.Header.BookingSourceRcd) && booking.Header.BookingSourceRcd?.ToUpper() == "TTY")
                                IsTTYMessage = true;
                        }
                        if (rsPassenger != null && rsPassenger.RecordCount > 0)
                        {
                            booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
                        }
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                      
                        }
                        if (rsRemark != null && rsRemark.RecordCount > 0)
                        {
                            booking.Remarks = booking.Remarks.FillBooking(rsRemark);
                        }
                        if (rsFees != null && rsFees.RecordCount > 0)
                        {
                            booking.Fees = booking.Fees.FillBooking(rsFees);
                        }
                        if (rsSegment != null && rsSegment.RecordCount > 0)
                        {
                            booking.Segments = booking.Segments.FillBooking(rsSegment);
                            numberOfPassenger = (short)booking.Segments[0].NumberOfUnits;
                        }
                        if (rsService != null && rsService.RecordCount > 0)
                        {
                            booking.Services = booking.Services.FillBooking(rsService);
                        }

                        // fill tax mapping 
                        if (taxes != null && taxes.Count > 0 && mappings != null && mappings.Count > 0)
                        {
                            mappings = taxes.FillTaxMapping(mappings);
                        }
                           

                        //if chnge flight: cancel flight
                        if (segments != null && segments.Count > 0 && booking.Segments != null && booking.Segments.Count > 0)
                        {

                            //determine change or add
                            foreach (FlightSegment f in segments)
                            {
                                if (f.SegmentStatusRcd.ToUpper() == "XX" && f.ExchangedSegmentId != Guid.Empty)
                                {
                                    // key =old seg  value new seg
                                    SegmentIdMappingChangeFlight.Add(new KeyValuePair<string, string>(f.BookingSegmentId.ToString(), f.ExchangedSegmentId.ToRsString()));
                                    
                                    changedSegmentFlag = true;
                                }
                                else if (f.SegmentStatusRcd.ToUpper() == "NN" && f.ExchangedSegmentId == Guid.Empty)
                                {
                                    alNewSegment.Add(f);
                                }
                            }

                            // remove change from alNewSegment
                            for (int i = 0; i < alNewSegment.Count; i++)
                            {
                                foreach (var element in SegmentIdMappingChangeFlight)
                                {
                                    FlightSegment ff = new FlightSegment();
                                    ff = (FlightSegment)alNewSegment[i];

                                    if(ff.BookingSegmentId == new Guid(element.Value))
                                    {
                                        alNewSegment.Remove(ff);
                                    }
                                }
                            }

                            if (alNewSegment != null && alNewSegment.Count > 0)
                            {
                                addSegmentFlag = true;
                            }

                            // for change
                            if (changedSegmentFlag && SegmentIdMappingChangeFlight != null && SegmentIdMappingChangeFlight.Count > 0)
                            {
                                // 
                                IsProcessSuccess = UpdatedSegmentAndMapping(booking, SegmentIdMappingChangeFlight, segments, mappings,
                                    strBookingId, strUserId, agencyCode, ref rsSegment, ref rsPayment, ref rsMapping,
                                    ref rsService, ref rsTax, ref rsQuote);
                                
                                // if change flight in same root: move SSR also
                                if (IsProcessSuccess && booking.Services != null && booking.Services.Count > 0)
                                {
                                    services = moveSSR(booking, SegmentIdMappingChangeFlight, ref rsService, segments, services);
                                }
                            }

                            // add SSR
                            if(services != null && services.Count > 0)
                            {
                                foreach(PassengerService ss in services)
                                {
                                   if(booking.FoundPassengerInfant(ss.PassengerId))
                                   {
                                       foundInfant = true;
                                   }
                                }
                            }

                            //for add
                            if (addSegmentFlag)
                            {
                                // for agg new segment
                                AddNewSegmentMapping(booking, alNewSegment,mappings,agencyCode, strUserId);
                            }
                        }
                        //end change flight

                        // update seat in old RS mapping
                        IList<entity.Mapping> rqstMappings = null;
                        IList<entity.PassengerService> listOfRqst  = null;

                        if (IsProcessSuccess && mappings != null && mappings.Count > 0)
                        {
                            rqstMappings = UpdatedSeatToMapping(mappings,ref rsMapping, strUserId);

                            if (rqstMappings != null && rqstMappings.Count > 0)
                            {
                                foreach (var m in rqstMappings)
                                {
                                    foreach (var s in booking.Segments)
                                    {
                                        if (s.BookingSegmentId == m.BookingSegmentId)
                                        {
                                            m.NumberOfUnits = s.NumberOfUnits;
                                        }
                                    }
                                }
                            }
                        }


                        //EDWDEV-771     
                        // do creat RQST SSR for TTY
                        if (IsTTYMessage && rqstMappings != null && rqstMappings.Count > 0 && booking.Mappings.Count > 0)
                        {
                            int CountPAXRQSTRequest = rqstMappings.Count;
                            int CountPAXRQSTProcess = 0;
                            bool isChangeSeatNullPassengerSwitchSeatLastPax = false;
                            bool isUseRQSTTogether = IsUseRQSTTogether(rqstMappings[0].BookingSegmentId, ref rsService);
                            bool isNullPassengerOUT_Change = CaseNullPassengerOUT_Change(rqstMappings) && isUseRQSTTogether;
                            bool isCaseNullPassengerChange_last_OUT = CaseNullPassengerChange_OUT(rqstMappings);

                            bool isFindRQST = IsFindRQST(rqstMappings, booking.Mappings, ref rsService);

                            if (isFindRQST == false)
                            {
                                // do NOT found RQST
                                foreach(var rqst in rqstMappings)
                                {
                                    List<entity.Mapping> RqRQSTOneByOne = new List<Mapping>();
                                    RqRQSTOneByOne.Add(rqst);

                                    // CancelSeatRQST regular
                                    listOfRqst = CreateSeatRQST(RqRQSTOneByOne, 1);
                                    listOfRqst.ToRecordset(ref rsService);
                                }
                                // END
                            }
                            else
                            {// do found RQST 

                                foreach (var rqstMapping in rqstMappings)
                                {
                                    CountPAXRQSTProcess = CountPAXRQSTProcess + 1;
                                    if (CountPAXRQSTProcess == CountPAXRQSTRequest)
                                    {
                                        isChangeSeatNullPassengerSwitchSeatLastPax = true;
                                    }

                                    List<entity.Mapping> RqRQSTOneByOne = new List<Mapping>();
                                    RqRQSTOneByOne.Add(rqstMapping);
                                   // Guid pid = rqstMapping.PassengerId;

                                    //check case passenger null or not
                                    bool isPassengerNull;

                                    isPassengerNull = IsPassengerNull(RqRQSTOneByOne, ref rsService);

                                    //check remove seat or not
                                    bool isRemoveSeat = IsRemoveSeat(RqRQSTOneByOne);

                                    bool isRemoveSeatAll = IsRemoveSeatAll(rqstMappings);

                                    bool isNullPassengerChangeSeatMoreOne = CaseNullPassengerChangeSeatMoreOne(rqstMappings);
                                    if (isNullPassengerChangeSeatMoreOne && isUseRQSTTogether)
                                        isPassengerNull = true;

                                    // bool IsRemoveSeat 
                                    if (isPassengerNull == true)
                                    {
                                        if (isRemoveSeat == false) // Null change 
                                        {

                                            { // Null  change 

                                                //update ASVC KD all passengers 
                                                UpdateStatusKDASCVPassengerNull(RqRQSTOneByOne, ref rsService);

                                                //cancel RQST XX that PAX NULL 
                                                CancelSeatRQSTNullPassenger(RqRQSTOneByOne, ref rsService, strUserId);

                                                //Create RQST HK to All PAX
                                                listOfRqst = CreateSeatRQSTForPassengerNull(RqRQSTOneByOne, ref rsService, numberOfPassenger, isNullPassengerChangeSeatMoreOne, isChangeSeatNullPassengerSwitchSeatLastPax);

                                                //find seat Characteristic
                                                List<SegmentSeatInfo> listOfSeatCharacteristic = AddSeatCharacteristicForPassNullWithRestPassenger(RqRQSTOneByOne, booking.Mappings);

                                                //est seat Characteristic
                                                if (isNullPassengerChangeSeatMoreOne == false)
                                                {
                                                    SetSeatCharacteristic(ref listOfRqst, listOfSeatCharacteristic);
                                                }
                                                else
                                                {
                                                    List<SegmentSeatInfo> listOfSeatCharacteristic2 = AddSeatCharacteristicForPassNullWithRestPassenger2(rqstMappings);

                                                    List<Mapping> theRestPassengerIds = new List<Mapping>();

                                                    foreach (var oldM in booking.Mappings)
                                                    {
                                                        bool exists = false;

                                                        foreach (var rq in rqstMappings)
                                                        {
                                                            if (rq.PassengerId == oldM.PassengerId)
                                                            {
                                                                exists = true;
                                                                break;
                                                            }
                                                        }

                                                        if (!exists)
                                                        {
                                                            theRestPassengerIds.Add(oldM);
                                                        }
                                                    }

                                                    List<SegmentSeatInfo> listOfSeatCharacteristic3 = new List<SegmentSeatInfo>();
                                                    if (theRestPassengerIds != null && theRestPassengerIds.Count > 0)
                                                    {
                                                        listOfSeatCharacteristic3 = AddSeatCharacteristicForPassNullWithRestPassenger2(theRestPassengerIds);
                                                    }


                                                    if (listOfRqst != null && listOfRqst.Count > 0)
                                                    {
                                                        foreach (var seat in listOfSeatCharacteristic2)
                                                        {
                                                            foreach (var ps in listOfRqst)
                                                            {
                                                                if (ps.BookingSegmentId == seat.BookingSegmentId)
                                                                {
                                                                    ps.ServiceText = ps.ServiceText + seat.SeatDetails;
                                                                }

                                                            }

                                                        }
                                                    }

                                                    // add listOfSeatCharacteristic3
                                                    if (listOfRqst != null && listOfRqst.Count > 0)
                                                    {
                                                        foreach (var seat in listOfSeatCharacteristic3)
                                                        {
                                                            foreach (var ps in listOfRqst)
                                                            {
                                                                if (ps.BookingSegmentId == seat.BookingSegmentId)
                                                                {
                                                                    ps.ServiceText = ps.ServiceText + seat.SeatDetails;
                                                                }

                                                            }

                                                        }
                                                    }
                                                }

                                                listOfRqst.ToRecordset(ref rsService);

                                                //for test
                                                //entity.Booking bb = new Booking();
                                                //List<entity.PassengerService> servicesTmpp = new List<PassengerService>();
                                                //bb.Services = servicesTmpp.FillBooking(rsService);
                                            }
                                        }
                                        else // case out seat
                                        {
                                            //do case passenger null and out seat
                                            //ceate RQST all passengers in segment WITHOUT paaseger that OUT
                                            listOfRqst = CreateRQSTSeparatePassengers(RqRQSTOneByOne, booking.Mappings);

                                            //NULL  first Change  last out 
                                            if (isCaseNullPassengerChange_last_OUT == true)
                                            {
                                                foreach (var rqst in rqstMappings)
                                                {
                                                    foreach (var r in listOfRqst)
                                                    {
                                                        if (rqst.BookingSegmentId == r.BookingSegmentId && rqst.PassengerId == r.PassengerId)
                                                        {
                                                            string seatCharacteristic = GetSeatCharacteristic(rqst.FlightId.ToString(), rqst.OriginRcd, 
                                                                rqst.DestinationRcd,rqst.BoardingClassRcd, rqst.SeatRow, rqst.SeatColumn);

                                                            r.ServiceText = string.Concat(rqst.SeatRow, rqst.SeatColumn,seatCharacteristic);
                                                        }
                                                    }
                                                }
                                            }
                                            listOfRqst.ToRecordset(ref rsService);

                                            //update ASVC KD all passengers in segment WITHOUT paaseger that OUT
                                            UpdateStatusKDRSeparatePassengers(RqRQSTOneByOne, booking.Mappings, ref rsService, strUserId);

                                            //cancel RQST that pax id is NULL   
                                            //CancelRQSTSeparatePassengers(RqRQSTOneByOne, booking.Mappings, ref rsService, strUserId);
                                            //cancel RQST XX that PAX NULL 
                                            CancelSeatRQSTNullPassenger(RqRQSTOneByOne, ref rsService, strUserId);

                                            //cancel ASVC XX that pax OUT
                                            CancelASVCSeparatePassengers(RqRQSTOneByOne, booking.Mappings, ref rsService, strUserId);
                                           
                                        }

                                    }
                                    else  //Normal passenger
                                    {
                                        if (isRemoveSeat == false) // change seat
                                        {
                                            UpdateStatusKDtoASVC(RqRQSTOneByOne, ref rsService);

                                            CancelSeatRQSTNormal(RqRQSTOneByOne, ref rsService, strUserId);

                                            if (isNullPassengerOUT_Change == false)
                                            {
                                                // CancelSeatRQST regular
                                                listOfRqst = CreateSeatRQST(RqRQSTOneByOne, 1);
                                                listOfRqst.ToRecordset(ref rsService);
                                            }
                                            else
                                            {
                                                // do case  isNullPassengerOUT_Change 
                                                listOfRqst = CreateSeatRQSTNullPassenger_OUT_Change(RqRQSTOneByOne, 1, ref rsService);
                                                listOfRqst.ToRecordset(ref rsService);

                                                //entity.Booking b2 = new Booking();
                                                //List<entity.PassengerService> servicesTmp2 = new List<PassengerService>();
                                                //b2.Services = servicesTmp2.FillBooking(rsService);

                                            }

                                        }
                                        else
                                        {
                                            // case remove all seat
                                            // for Normal both cancel case remove seat
                                            CancelASVC(RqRQSTOneByOne, ref rsService);

                                            // CancelSeatRQST regular for remove all                                  
                                            if (isRemoveSeatAll == true)
                                            {
                                                CancelSeatRQSTForRemoveSeatAllOfPassengerNull(RqRQSTOneByOne, ref rsService, strUserId);
                                            }
                                            else // CancelSeatRQST regular for remove some seat   
                                            {
                                                CancelSeatRQSTNormal(RqRQSTOneByOne, ref rsService, strUserId);
                                            }
                                        }
                                    }

                                }
                            }
                        
                            IsRQSTSeat = true;

                            //for test
                            entity.Booking b = new Booking();
                            List<entity.PassengerService> servicesTmp = new List<PassengerService>();
                            b.Services = servicesTmp.FillBooking(rsService);
                        }


                        // add segment && Mapping to RS
                        if (changedSegmentFlag)
                        {
                            booking.Mappings.ToRecordset(ref rsMapping);
                            booking.Segments.ToRecordset(ref rsSegment,true);
                        }

                        if (addSegmentFlag)
                        {
                            mappings.ToRecordset(ref rsMapping);
                            booking.Segments.ToRecordset(ref rsSegment);
                        }

                        //set SSR to RS
                        if (services == null)
                            services = new List<PassengerService>();

                        if (services.Count > 0)
                        {
                            // add ssr to RS
                            services.ToRecordset(ref rsService);
                        }

                        if (taxes == null)
                            taxes = new List<Tax>();

                        if (taxes.Count > 0)
                        {
                            // add tax to RS
                            taxes.ToRecordset(ref rsTax);
                        }

                        //Add fee to RS
                        if (fees != null && fees.Count > 0)
                        {
                            // map fee to rs
                            fees.ToRecordset(ref rsFees);
                        }

                        // set pnrupd create by
                        if (IsProcessSuccess)
                        {

                            #region not use

                            // UpdateFeeCreateBy(ref rsFees, strUserId);
                            //   UpdateMappingCreateBy(ref rsMapping, strUserId);
                            // UpdateSegmentCreateBy(ref rsSegment, strUserId);
                            // UpdateSSRCreateBy(ref rsService, strUserId);

                            /*
                            // cancel the existing one RQST both remove seat or change seat
                            if (rsService != null) 
                            {                              
                                if (isPassengerNull == true )
                                {
                                    if(IsRemoveSeat == true)
                                    {
                                        // remove seat for null passenger
                                        
                                    }
                                    else
                                    {
                                        //change seat for null passenger
                                       // CancelSeatRQST(rqstMappings, ref rsService, strUserId);
                                    }
                                    
                                }
                                else
                                {
                                    // CancelSeatRQST regular
                                   // CancelSeatRQST(rqstMappings, ref rsService, strUserId);
                                }
                                   
                            }
                            */

                            #endregion
                        }

                        // test assign to old mapping
                        UpdatedRSOldMapping(booking, SegmentIdMappingChangeFlight, segments, mappings,
                                    strBookingId, strUserId, agencyCode, ref rsSegment, ref rsPayment, ref rsMapping,
                                    ref rsService, ref rsTax, ref rsQuote);



                        if (!booking.IsLockBooking())
                        {
                            if (!foundInfant)
                            {
                                // is every thing ok allwed to save booking
                                if (IsProcessSuccess)
                                {
                                    //enumSaveError = objBooking.Save(ref rsHeader,
                                    //                                  ref rsSegment,
                                    //                                  ref rsPassenger,
                                    //                                  ref rsRemark,
                                    //                                  ref rsPayment,
                                    //                                  ref rsMapping,
                                    //                                  ref rsService,
                                    //                                  ref rsTax,
                                    //                                  ref rsFees,
                                    //                                  ref createTickets,
                                    //                                  ref readBooking,
                                    //                                  ref readOnly,
                                    //                                  ref bSetLock,
                                    //                                  ref bCheckSeatAssignment,
                                    //                                  ref bCheckSessionTimeOut);

                                    if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                                    {
                                        //EDWDEV-771  call message SP:LogTeletypeMessage
                                        if(IsTTYMessage && IsRQSTSeat)
                                            objBooking.BookingChangeQueueSave(ref strBookingId, ref strUserId);

                                        //clear RS
                                        RecordsetHelper.ClearRecordset(ref rsHeader);
                                        RecordsetHelper.ClearRecordset(ref rsSegment);
                                        RecordsetHelper.ClearRecordset(ref rsPassenger);
                                        RecordsetHelper.ClearRecordset(ref rsRemark);
                                        RecordsetHelper.ClearRecordset(ref rsPayment);
                                        RecordsetHelper.ClearRecordset(ref rsMapping);
                                        RecordsetHelper.ClearRecordset(ref rsService);
                                        RecordsetHelper.ClearRecordset(ref rsTax);
                                        RecordsetHelper.ClearRecordset(ref rsFees);
                                        RecordsetHelper.ClearRecordset(ref rsQuote);

                                        //read booking  
                                        readSuccess = readForSave(strBookingId, strUserId, ref rsHeader, ref rsSegment, ref rsPassenger, ref rsRemark, ref rsPayment, ref rsMapping, ref rsService, ref rsTax, ref rsFees, ref rsQuote);

                                        if (readSuccess)
                                        {
                                            // convert rs to booking for preparing header and mapping  value
                                            if (rsHeader != null && rsHeader.RecordCount > 0)
                                            {
                                                bookingResponse.Header = booking.Header.FillBooking(rsHeader);
                                            }
                                            if (rsSegment != null && rsSegment.RecordCount > 0)
                                            {
                                                bookingResponse.Segments = booking.Segments.FillBooking(rsSegment);
                                            }
                                            if (rsPassenger != null && rsPassenger.RecordCount > 0)
                                            {
                                                bookingResponse.Passengers = booking.Passengers.FillBooking(rsPassenger);
                                            }
                                            if (rsFees != null && rsFees.RecordCount > 0)
                                            {
                                                bookingResponse.Fees = booking.Fees.FillBooking(rsFees);
                                            }
                                            if (rsRemark != null && rsRemark.RecordCount > 0)
                                            {
                                                bookingResponse.Remarks = booking.Remarks.FillBooking(rsRemark);
                                            }
                                            if (rsQuote != null && rsQuote.RecordCount > 0)
                                            {
                                                bookingResponse.Quotes = booking.Quotes.FillBooking(rsQuote);
                                            }
                                            if (rsTax != null && rsTax.RecordCount > 0)
                                            {
                                                bookingResponse.Taxs = booking.Taxs.FillBooking(rsTax);
                                            }
                                            if (rsService != null && rsService.RecordCount > 0)
                                            {
                                                bookingResponse.Services = booking.Services.FillBooking(rsService);
                                            }
                                            if (rsMapping != null && rsMapping.RecordCount > 0)
                                            {
                                                bookingResponse.Mappings = booking.Mappings.FillBooking(rsMapping);
                                            }
                                            if (rsPayment != null && rsPayment.RecordCount > 0)
                                            {
                                                bookingResponse.Payments = booking.Payments.FillBooking(rsPayment);
                                            }
                                        }
                                        else
                                        {
                                            strError = "Read booking fail";
                                            throw new ModifyBookingException("Read booking fail", "B008");
                                        }
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                                    {
                                        strError = "BOOKINGINUSE";
                                        throw new BookingSaveException("BOOKINGINUSE");
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                                    {
                                        strError = "BOOKINGREADERROR";
                                        throw new BookingSaveException("BOOKINGREADERROR");
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                                    {
                                        strError = "DATABASEACCESS";
                                        throw new BookingSaveException("DATABASEACCESS");
                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                                    {
                                        strError = "DUPLICATESEAT";
                                        throw new BookingSaveException("Duplicated Seat", "277");

                                    }
                                    else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                                    {
                                        strError = "SESSIONTIMEOUT";
                                        throw new BookingSaveException("SESSIONTIMEOUT", "H005");
                                    }
                                }
                                // end process
                                else
                                {
                                    throw new ModifyBookingException("Modify booking fail", "300");
                                }
                            }
                            else
                            {
                                throw new ModifyBookingException("Can not add SSR with Infant.", "L008");
                            }
                        }
                        else
                        {
                            throw new ModifyBookingException("Booking is Locked.", "L001");
                        }
                    }
                    else
                    {
                        throw new ModifyBookingException("Read booking fail", "B008");
                    }
                }
            }
            catch (ModifyBookingException ex)
            {
                if (ex.ErrorCode == "L008")
                    throw new ModifyBookingException("Infant is not allowed to book SSR.", "L008");
                else if(ex.ErrorCode == "L001")
                    throw new ModifyBookingException("Booking is Locked.", "L001");
                else
                    throw new ModifyBookingException(ex.Message, "");
            }
            catch (BookingSaveException)
            {
                throw new ModifyBookingException(strError, "");
            }
            catch
            {
                throw;// new ModifyBookingException(strError);
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);

            }
            return bookingResponse;
        }

        private bool UpdatedRSOldMapping(Booking booking, List<KeyValuePair<string, string>> SegmentIdMappingChangeFlight,
                                IList<entity.FlightSegment> segments,
                                IList<entity.Mapping> mappings,
                                string strBookingId, string strUserId, string agencyCode,
                                ref ADODB.Recordset rsSegment,
                                ref ADODB.Recordset rsPayment,
                                 ref ADODB.Recordset rsMapping,
                                 ref ADODB.Recordset rsService,
                                 ref ADODB.Recordset rsTax,
                                ref ADODB.Recordset rsQuote
                                )
        {
            bool IsProcessSuccess = false;
            foreach (var element in SegmentIdMappingChangeFlight)
            {
                foreach (FlightSegment f in segments)
                {
                    foreach (Mapping newMapping in mappings)
                    {
                        foreach (Mapping oldMapping in booking.Mappings)
                        {
                            //if NN                             new seg = value                              value = new mapping                                              key = old mapping
                            if (f.SegmentStatusRcd == "NN" && f.BookingSegmentId == new Guid(element.Value) && new Guid(element.Value) == newMapping.BookingSegmentId && new Guid(element.Key) == oldMapping.BookingSegmentId && newMapping.PassengerId == oldMapping.PassengerId)
                            {
                                if (rsMapping != null && rsMapping.RecordCount > 0)
                                {
                                    rsMapping.MoveFirst();
                                    while (!rsMapping.EOF)
                                    {
                                        if (RecordsetHelper.ToString(rsMapping, "passenger_status_rcd") == "XX" && RecordsetHelper.ToGuid(rsMapping, "booking_segment_id") == oldMapping.BookingSegmentId)
                                        {
                                            rsMapping.Fields["e_ticket_status"].Value = "E";
                                            rsMapping.Fields["exchanged_segment_id"].Value = f.BookingSegmentId.ToRsString();

                                        }
                                        rsMapping.MoveNext();
                                    }
                                }


                            }

                        }
                    }
                }
            }

            return IsProcessSuccess;

        }




        private bool UpdatedSegmentAndMapping(Booking booking, List<KeyValuePair<string, string>> SegmentIdMappingChangeFlight, 
                                IList<entity.FlightSegment> segments, 
                                IList<entity.Mapping> mappings,
                                string strBookingId, string strUserId, string agencyCode,
                                ref  ADODB.Recordset rsSegment,
                                ref  ADODB.Recordset rsPayment,
                                 ref  ADODB.Recordset rsMapping,
                                 ref  ADODB.Recordset rsService,
                                 ref  ADODB.Recordset rsTax,
                                ref  ADODB.Recordset rsQuote
                                )
        {
            bool IsProcessSuccess = false;
            foreach (var element in SegmentIdMappingChangeFlight)
            {
                foreach (FlightSegment f in segments)
                {
                    foreach (Mapping newMapping in mappings)
                    {
                        foreach (Mapping oldMapping in booking.Mappings)
                        {
                            //if NN                             new seg = value                              value = new mapping                                              key = old mapping
                            if (f.SegmentStatusRcd == "NN" && f.BookingSegmentId == new Guid(element.Value) && new Guid(element.Value) == newMapping.BookingSegmentId && new Guid(element.Key) == oldMapping.BookingSegmentId && newMapping.PassengerId == oldMapping.PassengerId)
                            {
                                // cancel seg first
                                IsProcessSuccess = CancelSegment(strBookingId, strUserId, agencyCode, ref rsSegment, ref rsPayment, ref rsMapping, ref rsService, ref rsTax, ref rsQuote, element.Key.ToUpper());

                                //change : add new NN
                                if (IsProcessSuccess)
                                {
                                    //assign booking id to new segment
                                    f.BookingId = booking.Header.BookingId;

                                    // assign new seg id to ExchangedSegmentId in mapping
                                    oldMapping.ExchangedSegmentId = new Guid(element.Value);
                                    oldMapping.ExchangedDateTime = DateTime.Now;

                                    //fix change flight e ticket status old mapping = E
                                  //  oldMapping.ETicketStatus = "E";


                                    //assign new mapping
                                    newMapping.Firstname = oldMapping.Firstname;
                                    newMapping.Lastname = oldMapping.Lastname;
                                    newMapping.PassengerTypeRcd = oldMapping.PassengerTypeRcd;
                                    newMapping.TitleRcd = oldMapping.TitleRcd;
                                    newMapping.PassengerStatusRcd = "OK";
                                    newMapping.GenderTypeRcd = oldMapping.GenderTypeRcd;

                                    newMapping.AirlineRcd = f.AirlineRcd;
                                    newMapping.FlightNumber = f.FlightNumber;
                                    newMapping.FlightId = f.FlightId;
                                    newMapping.FlightConnectionId = f.FlightConnectionId;
                                    newMapping.DepartureDate = f.DepartureDate;
                                    newMapping.DepartureTime = f.DepartureTime;
                                    newMapping.ArrivalDate = f.ArrivalDate;
                                    newMapping.ArrivalTime = f.ArrivalTime;
                                    newMapping.JourneyTime = f.JourneyTime;
                                    newMapping.OriginRcd = f.OriginRcd;
                                    newMapping.DestinationRcd = f.DestinationRcd;
                                    newMapping.OdOriginRcd = f.OdOriginRcd;
                                    newMapping.OdDestinationRcd = f.OdDestinationRcd;
                                    newMapping.BookingClassRcd = f.BookingClassRcd;
                                    newMapping.BoardingClassRcd = f.BoardingClassRcd;
                                    newMapping.AgencyCode = agencyCode;

                                    //fix change flight e ticket status new mapping = O
                                    //newMapping.ETicketStatus = "O";
                                    //newMapping.YqAmount = 44;
                                    //newMapping.YqAmountIncl = 44;
                                    //newMapping.TaxAmount = 35;
                                    //newMapping.TaxAmountIncl = 35;

                                    newMapping.FareId = f.FareId;

                                    newMapping.UpdateBy = new Guid(strUserId);
                                    newMapping.UpdateDateTime = DateTime.Now;
                                    newMapping.CreateBy = new Guid(strUserId);
                                    newMapping.CreateDateTime = DateTime.Now;


                                }
                                // add new seg: NN
                                if (f.SegmentStatusRcd == "NN")
                                    booking.Segments.Add(f);
                            }

                        }
                        // add New Mapping to booking
                        if (booking.Mappings == null)
                            booking.Mappings = new List<Mapping>();

                        if (newMapping.Firstname != null)
                            booking.Mappings.Add(newMapping);

                    }
                }
            }

            return IsProcessSuccess;

        }


        private void AddNewSegmentMapping(Booking booking, ArrayList alNewSegment, IList<Mapping> mappings, string agencyCode,string strUserId)
        {
            for (int i = 0; i < alNewSegment.Count; i++)
            {
                FlightSegment newSegment = new FlightSegment();
                newSegment = (FlightSegment)alNewSegment[i];

                if (newSegment.SegmentStatusRcd.ToUpper() == "NN")
                {
                    newSegment.BookingId = booking.Header.BookingId;

                    // add NN to booking 
                    booking.Segments.Add(newSegment);

                    // add mapping
                    foreach (Mapping newMapping in mappings)
                    {
                        foreach (Mapping oldMapping in booking.Mappings)
                        {
                            if (newMapping.PassengerId == oldMapping.PassengerId && newMapping.BookingSegmentId == newSegment.BookingSegmentId)
                            {
                                newMapping.Firstname = oldMapping.Firstname;
                                newMapping.Lastname = oldMapping.Lastname;
                                newMapping.PassengerTypeRcd = oldMapping.PassengerTypeRcd;
                                newMapping.PassengerStatusRcd = "OK";

                                newMapping.TitleRcd = oldMapping.TitleRcd;
                                newMapping.GenderTypeRcd = oldMapping.GenderTypeRcd;


                                newMapping.AirlineRcd = newSegment.AirlineRcd;
                                newMapping.FlightNumber = newSegment.FlightNumber;
                                newMapping.FlightId = newSegment.FlightId;
                                newMapping.FlightConnectionId = newSegment.FlightConnectionId;
                                newMapping.DepartureDate = newSegment.DepartureDate;
                                newMapping.DepartureTime = newSegment.DepartureTime;
                                newMapping.ArrivalDate = newSegment.ArrivalDate;
                                newMapping.ArrivalTime = newSegment.ArrivalTime;
                                newMapping.JourneyTime = newSegment.JourneyTime;
                                newMapping.OriginRcd = newSegment.OriginRcd;
                                newMapping.DestinationRcd = newSegment.DestinationRcd;
                                newMapping.OdOriginRcd = newSegment.OdOriginRcd;
                                newMapping.OdDestinationRcd = newSegment.OdDestinationRcd;
                                newMapping.BookingClassRcd = newSegment.BookingClassRcd;
                                newMapping.BoardingClassRcd = newSegment.BoardingClassRcd;
                                newMapping.AgencyCode = agencyCode;

                                newMapping.UpdateBy = new Guid(strUserId);
                                newMapping.UpdateDateTime = DateTime.Now;
                                newMapping.CreateBy = new Guid(strUserId);
                                newMapping.CreateDateTime = DateTime.Now;

                            }

                        }

                    }
                }

            }

        }

        private void SetNameToRsPassenger(ref ADODB.Recordset rsPassenger, IList<Entity.NameChange> NameChange, string strUserId)
        {
            foreach (NameChange n in NameChange)
            {
                rsPassenger.MoveFirst();
                while (!rsPassenger.EOF)
                {
                    if (n.PassengerId == RecordsetHelper.ToGuid(rsPassenger, "passenger_id"))
                    {
                        if (!string.IsNullOrEmpty(n.TitleRcd))
                        {
                            rsPassenger.Fields["title_rcd"].Value = n.TitleRcd;
                        }
                        if (!string.IsNullOrEmpty(n.GenderTypeRcd))
                        {
                            rsPassenger.Fields["gender_type_rcd"].Value = n.GenderTypeRcd.ToUpper();
                        }
                        if (!string.IsNullOrEmpty(n.Firstname))
                            rsPassenger.Fields["firstname"].Value = n.Firstname.ToUpper();
                        if (!string.IsNullOrEmpty(n.Middlename))
                            rsPassenger.Fields["middlename"].Value = n.Middlename.ToUpper();
                        if (!string.IsNullOrEmpty(n.Lastname))
                            rsPassenger.Fields["lastname"].Value = n.Lastname.ToUpper();
                        if (n.DateOfBirth != DateTime.MinValue)
                            rsPassenger.Fields["date_of_birth"].Value = n.DateOfBirth;

                        rsPassenger.Fields["update_date_time"].Value = DateTime.Now;
                        rsPassenger.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                        rsPassenger.Fields["create_date_time"].Value = DateTime.Now;
                        rsPassenger.Fields["create_by"].Value = new Guid(strUserId).ToRsString();
                    }
                    rsPassenger.MoveNext();
                }
            }

        }

        private void SetNameToRsMapping(ref ADODB.Recordset rsMapping, IList<Entity.NameChange> NameChange, string strUserId)
        {
            foreach (NameChange n in NameChange)
            {
                rsMapping.MoveFirst();
                while (!rsMapping.EOF)
                {
                    if (n.PassengerId == RecordsetHelper.ToGuid(rsMapping, "passenger_id"))
                    {
                        if (!string.IsNullOrEmpty(n.TitleRcd))
                        {
                            rsMapping.Fields["title_rcd"].Value = n.TitleRcd.ToUpper();
                        }
                        if (!string.IsNullOrEmpty(n.GenderTypeRcd))
                        {
                            rsMapping.Fields["gender_type_rcd"].Value = n.GenderTypeRcd.ToUpper();
                        }
                        if (!string.IsNullOrEmpty(n.Firstname))
                            rsMapping.Fields["firstname"].Value = n.Firstname.ToUpper();
                        if (!string.IsNullOrEmpty(n.Lastname))
                            rsMapping.Fields["lastname"].Value = n.Lastname.ToUpper();
                        if (n.DateOfBirth != DateTime.MinValue)
                            rsMapping.Fields["date_of_birth"].Value = n.DateOfBirth;

                        rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                        rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                        rsMapping.Fields["create_date_time"].Value = DateTime.Now;
                        rsMapping.Fields["create_by"].Value = new Guid(strUserId).ToRsString();
                    }
                    rsMapping.MoveNext();
                }
            }

        }
        private void SetSeatToSameRsMappingX(ref ADODB.Recordset rsMapping, IList<Entity.Booking.Mapping> Mappings, string strUserId)
        {
            foreach (entity.Mapping m in Mappings)
            {
                rsMapping.MoveFirst();
                while (!rsMapping.EOF)
                {
                    if (RecordsetHelper.ToGuid(rsMapping, "booking_segment_id") == m.BookingSegmentId && RecordsetHelper.ToGuid(rsMapping, "passenger_id") == m.PassengerId && m.SeatNumber.Length > 0)
                    {
                        rsMapping.Fields["seat_column"].Value = m.SeatColumn;
                        rsMapping.Fields["seat_row"].Value = m.SeatRow;
                        rsMapping.Fields["seat_number"].Value = m.SeatNumber;

                        if (m.SeatFeeRcd != null)
                            rsMapping.Fields["seat_fee_rcd"].Value = m.SeatFeeRcd;

                        rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                        rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                        rsMapping.Fields["create_date_time"].Value = DateTime.Now;
                        rsMapping.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                        break;
                    }
                    rsMapping.MoveNext();
                }
            }

        }

        private void SetSeatToSameRsMapping(ref ADODB.Recordset rsMapping, Entity.Booking.Mapping m, string strUserId)
        {
            rsMapping.MoveFirst();
            while (!rsMapping.EOF)
            {
                if (RecordsetHelper.ToGuid(rsMapping, "booking_segment_id") == m.BookingSegmentId && RecordsetHelper.ToGuid(rsMapping, "passenger_id") == m.PassengerId && m.SeatNumber.Length > 0)
                {
                    rsMapping.Fields["seat_column"].Value = m.SeatColumn;
                    rsMapping.Fields["seat_row"].Value = m.SeatRow;
                    rsMapping.Fields["seat_number"].Value = m.SeatNumber;

                    if (m.SeatFeeRcd != null)
                        rsMapping.Fields["seat_fee_rcd"].Value = m.SeatFeeRcd;

                    rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                    rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                    rsMapping.Fields["create_date_time"].Value = DateTime.Now;
                    rsMapping.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                    break;
                }
                rsMapping.MoveNext();
            }
        }

        private void SetSeatToNewRsMappingX(ref ADODB.Recordset rsMapping, List<KeyValuePair<string, string>> objSegmentIdMapping, IList<Entity.Booking.Mapping> Mappings, string strUserId)
        {
            foreach (var element in objSegmentIdMapping)
            {
                foreach (entity.Mapping m in Mappings)
                {
                    rsMapping.MoveFirst();
                    while (!rsMapping.EOF)
                    {
                        //old id == mapping_Segid && RSSegId == new id && RS_passId == mapping_PassId
                        if (element.Key.ToUpper().Equals(m.BookingSegmentId.ToString().ToUpper()) && RecordsetHelper.ToGuid(rsMapping, "booking_segment_id") == new Guid(element.Value) && RecordsetHelper.ToGuid(rsMapping, "passenger_id") == m.PassengerId && m.SeatNumber != null && m.SeatNumber.Length > 0)
                        {
                            rsMapping.Fields["seat_column"].Value = m.SeatColumn;
                            rsMapping.Fields["seat_row"].Value = m.SeatRow;
                            rsMapping.Fields["seat_number"].Value = m.SeatNumber;

                            if (m.SeatFeeRcd != null)
                                rsMapping.Fields["seat_fee_rcd"].Value = m.SeatFeeRcd;

                            rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                            rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                            rsMapping.Fields["create_date_time"].Value = DateTime.Now;
                            rsMapping.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                            break;
                        }

                        rsMapping.MoveNext();
                    }
                }
            }

        }

        private void SetSeatToNewRsMapping(ref ADODB.Recordset rsMapping, List<KeyValuePair<string, string>> objSegmentIdMapping, Entity.Booking.Mapping m, string strUserId)
        {
            foreach (var element in objSegmentIdMapping)
            {
                    rsMapping.MoveFirst();
                    while (!rsMapping.EOF)
                    {
                        //old id == mapping_Segid && RSSegId == new id && RS_passId == mapping_PassId
                        if (element.Key.ToUpper().Equals(m.BookingSegmentId.ToString().ToUpper()) && RecordsetHelper.ToGuid(rsMapping, "booking_segment_id") == new Guid(element.Value) && RecordsetHelper.ToGuid(rsMapping, "passenger_id") == m.PassengerId && m.SeatNumber != null && m.SeatNumber.Length > 0)
                        {
                            rsMapping.Fields["seat_column"].Value = m.SeatColumn;
                            rsMapping.Fields["seat_row"].Value = m.SeatRow;
                            rsMapping.Fields["seat_number"].Value = m.SeatNumber;

                            if (m.SeatFeeRcd != null)
                                rsMapping.Fields["seat_fee_rcd"].Value = m.SeatFeeRcd;

                            rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                            rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                            rsMapping.Fields["create_date_time"].Value = DateTime.Now;
                            rsMapping.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                            break;
                        }

                        rsMapping.MoveNext();
                    }
            }

        }

        private IList<Entity.Booking.Mapping> SetSeatToMapping(IList<Entity.Booking.Mapping> Mappings, IList<Entity.SeatAssign> seatAssigns, string strUserId)
        {
            for (int i = 0; i < seatAssigns.Count; i++)
            {
                foreach (Entity.Booking.Mapping m in Mappings)
                {
                    if (seatAssigns[i].BookingSegmentID.ToUpper().Equals(m.BookingSegmentId.ToString().ToUpper()) && seatAssigns[i].PassengerID.ToUpper().Equals(m.PassengerId.ToString().ToUpper()))
                    {
                        m.SeatColumn = seatAssigns[i].SeatColumn;
                        m.SeatRow = seatAssigns[i].SeatRow;

                        if (seatAssigns[i].SeatFeeRcd != null)
                            m.SeatFeeRcd = seatAssigns[i].SeatFeeRcd;

                        m.SeatNumber = seatAssigns[i].SeatNumber;

                        m.UpdateBy = new Guid(strUserId);
                        m.UpdateDateTime = DateTime.Now;
                    }
                }
            }

            return Mappings;
        }
       
        private IList<Entity.Booking.Mapping>  CancelSeatInChangeFlight(List<KeyValuePair<string, string>> objSegmentIdMapping, IList<Entity.Booking.Mapping> Mappings)
        {
            foreach (var element in objSegmentIdMapping)
            {
                foreach (entity.Mapping m in Mappings)
                {
                    if (element.Key.ToUpper().Equals(m.BookingSegmentId.ToString().ToUpper()) && m.SeatNumber != null && m.SeatNumber.Length > 0)
                    {
                        m.SeatColumn = string.Empty;
                        m.SeatFeeRcd = string.Empty;
                        m.SeatNumber = string.Empty;
                        m.SeatRow = 0;
                    }
                }
            }

            return Mappings;
        }

        private IList<Entity.Booking.Fee> SetNewIdToFee(List<KeyValuePair<string, string>> objSegmentIdMapping, IList<Entity.Booking.Fee> correctionFee)
        {
            foreach (var element in objSegmentIdMapping)
            {
                foreach (entity.Fee f in correctionFee)
                {
                    if (f.BookingSegmentId == new Guid(element.Key))
                    {
                        f.BookingSegmentId = new Guid(element.Value);
                    }
                }
            }

            return correctionFee;
        }

        private IList<Entity.Booking.Fee> AddCreateBy(IList<Entity.Booking.Fee> correctionFee, Guid user)
        {
            foreach (entity.Fee f in correctionFee)
            {
                f.CreateBy = user;
                f.CreateDateTime = DateTime.Now;
                f.UpdateBy = user;
                f.UpdateDateTime = DateTime.Now;
            }

            return correctionFee;
        }
        
        private void UpdateMappingCreateBy(ref ADODB.Recordset rsMapping, string strUserId)
        {
            if (rsMapping != null && rsMapping.RecordCount > 0)
            {
                rsMapping.MoveFirst();
                while (!rsMapping.EOF)
                {
                    if (RecordsetHelper.ToGuid(rsMapping, "create_by") == Guid.Empty)
                    {
                        rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                        rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                        rsMapping.Fields["create_date_time"].Value = DateTime.Now;
                        rsMapping.Fields["create_by"].Value = new Guid(strUserId).ToRsString();
                    }
                    rsMapping.MoveNext();
                }
            }
        }

        private void VoidSeatFee(ref ADODB.Recordset rsFees, IList<Entity.SeatAssign> seatAssigns, string strUserId)
        {
            foreach (Entity.SeatAssign f in seatAssigns)
            {
                rsFees.MoveFirst();
                while (!rsFees.EOF)
                {
                    if (RecordsetHelper.ToString(rsFees, "fee_category_rcd") == "SEAT"
                        && RecordsetHelper.ToGuid(rsFees, "account_fee_by") == Guid.Empty
                        && RecordsetHelper.ToDateTime(rsFees, "void_date_time") == DateTime.MinValue
                        && new Guid(f.PassengerID) == RecordsetHelper.ToGuid(rsFees, "passenger_id")
                        && new Guid(f.BookingSegmentID) == RecordsetHelper.ToGuid(rsFees, "booking_segment_id")
                        )
                    {
                        rsFees.Fields["void_date_time"].Value = DateTime.Now;
                        rsFees.Fields["void_by"].Value = new Guid(strUserId).ToRsString();
                    }
                    rsFees.MoveNext();
                }
            }

        }

        private void UpdateFeeCreateBy(ref ADODB.Recordset rsFees, string strUserId)
        {
            if (rsFees != null  && rsFees.RecordCount > 0)
            {
                rsFees.MoveFirst();
                while (!rsFees.EOF)
                {
                    if (RecordsetHelper.ToGuid(rsFees, "create_by") == Guid.Empty)
                    {
                        rsFees.Fields["update_date_time"].Value = DateTime.Now;
                        rsFees.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                        rsFees.Fields["create_date_time"].Value = DateTime.Now;
                        rsFees.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                       // rsFees.Fields["account_fee_date_time"].Value = DateTime.Now;
                      //  rsFees.Fields["account_fee_by"].Value = new Guid(strUserId).ToRsString();
                    }
                    rsFees.MoveNext();
                }
            }
        }
        private void UpdateSSRCreateBy(ref ADODB.Recordset rsService, string strUserId)
        {
            if (rsService != null && rsService.RecordCount > 0)
            {
                rsService.MoveFirst();
                while (!rsService.EOF)
                {
                    if (RecordsetHelper.ToGuid(rsService, "create_by") == Guid.Empty && rsService.Fields["special_service_rcd"].Value != "RQST")
                    {
                        rsService.Fields["update_date_time"].Value = DateTime.Now;
                        rsService.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                        rsService.Fields["create_date_time"].Value = DateTime.Now;
                        rsService.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                    }
                    rsService.MoveNext();
                }
            }
        }


        private void UpdateSegmentCreateBy(ref ADODB.Recordset rsSegment, string strUserId)
        {
            if (rsSegment != null && rsSegment.RecordCount > 0)
            {
                rsSegment.MoveFirst();
                while (!rsSegment.EOF)
                {
                    if (RecordsetHelper.ToGuid(rsSegment, "create_by") == Guid.Empty)
                    {
                        rsSegment.Fields["update_date_time"].Value = DateTime.Now;
                        rsSegment.Fields["update_by"].Value = new Guid(strUserId).ToRsString();

                        rsSegment.Fields["create_date_time"].Value = DateTime.Now;
                        rsSegment.Fields["create_by"].Value = new Guid(strUserId).ToRsString();

                    }
                    rsSegment.MoveNext();
                }
            }
        }

        private void UpdateHeader(ref ADODB.Recordset rsHeader, Entity.Booking.BookingHeader Header, string strUserId)
        {
            if (rsHeader != null && rsHeader.RecordCount > 0)
            {
                if (Header.BookingId == RecordsetHelper.ToGuid(rsHeader, "booking_id"))
                {
                    rsHeader.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                    rsHeader.Fields["update_date_time"].Value = DateTime.Now;
                }
            }
        }

        private IList<Remark> AddRemark(string strBookingId, string agencyCode, string strUserId)
        {
            IList<Remark> remarkList = new List<Remark>();
            Remark remark = new Remark();
            remark.BookingId = new Guid(strBookingId);
            remark.AddedBy = "API / " + agencyCode;
            remark.AgencyCode = agencyCode;
            remark.RemarkText = "Change Reference: API";
            remark.RemarkTypeRcd = "HISTORY";
            remark.UpdateBy = new Guid(strUserId);
            remark.UpdateDateTime = DateTime.Now; ;
            remark.CreateBy = new Guid(strUserId);
            remark.CreateDateTime = DateTime.Now; ;
            remarkList.Add(remark);
            return remarkList;
        }
        
        public bool UpdatePassengerInfo(string strBookingId,
                                            IList<Entity.Booking.Passenger> passenger,
                                            IList<Entity.Booking.Mapping> mapping,
                                            string strUserId,
                                            string agencyCode,
                                            bool bNoVat)
        {
            tikAeroProcess.Booking objBooking = null;
            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;

            string strXml = string.Empty;
            bool createTickets = false;
            bool readBooking = false;
            bool readOnly = false;
            bool bSetLock = false;
            bool bCheckSeatAssignment = false;
            bool bCheckSessionTimeOut = false;
            string paymentType = string.Empty;
            string language = "EN";
            string currency = string.Empty;
            entity.Booking bookingResponse = new entity.Booking();
            entity.Booking booking = new entity.Booking();
            var objSegmentIdMapping = new List<KeyValuePair<string, string>>();
            bool result = false;
            tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                }

                if (!String.IsNullOrEmpty(strBookingId))
                {
                    //read booking first for prepare RS 
                    if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                        string.Empty,
                        0,
                        ref rsHeader,
                        ref rsSegment,
                        ref rsPassenger,
                        ref rsRemark,
                        ref rsPayment,
                        ref rsMapping,
                        ref rsService,
                        ref rsTax,
                        ref rsFees,
                        ref rsQuote,
                        false,
                        false,
                        "{" + strUserId.ToUpper() + "}",
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        string.Empty,
                        string.Empty) == true)
                    {
                        // convert rs to booking for preparing header and mapping  value
                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);
                            currency = booking.Header.CurrencyRcd;
                            language = booking.Header.LanguageRcd;
                        }
                        if (rsPassenger != null && rsPassenger.RecordCount > 0)
                        {
                            booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
                        }
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        }
                    }

                    foreach (entity.Passenger reqPass in passenger)
                    {
                        foreach (entity.Passenger p in booking.Passengers)
                        {
                            if (reqPass.PassengerId == p.PassengerId)
                            {
                                p.DocumentTypeRcd = reqPass.DocumentTypeRcd;
                                p.DocumentNumber = reqPass.DocumentNumber;
                                p.PassportNumber = reqPass.PassportNumber;
                                p.PassportIssueDate = reqPass.PassportIssueDate;
                                p.PassportExpiryDate = reqPass.PassportExpiryDate;
                                p.PassportIssuePlace = reqPass.PassportIssuePlace;
                                p.PassportIssueCountryRcd = reqPass.PassportIssueCountryRcd;
                                p.NationalityRcd = reqPass.NationalityRcd;
                                p.PassportBirthPlace = reqPass.PassportBirthPlace;

                            }
                        }
                    }

                    foreach (entity.Passenger p in booking.Passengers)
                    {
                        rsPassenger.MoveFirst();
                        while (!rsPassenger.EOF)
                        {
                            if (p.PassengerId == RecordsetHelper.ToGuid(rsPassenger, "passenger_id"))
                            {
                                rsPassenger.Fields["document_type_rcd"].Value = p.DocumentTypeRcd;
                                rsPassenger.Fields["passport_birth_place"].Value = p.PassportBirthPlace;
                                rsPassenger.Fields["passport_number"].Value = p.PassportNumber;
                                rsPassenger.Fields["passport_issue_date"].Value = p.PassportIssueDate;
                                rsPassenger.Fields["passport_expiry_date"].Value = p.PassportExpiryDate;
                                rsPassenger.Fields["passport_issue_place"].Value = p.PassportIssuePlace;
                                rsPassenger.Fields["passport_issue_country_rcd"].Value = p.PassportIssueCountryRcd;
                                rsPassenger.Fields["nationality_rcd"].Value = p.NationalityRcd;
                                rsPassenger.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                                break;
                            }
                            rsPassenger.MoveNext();
                        }
                    }


                    // save booking
                    enumSaveError = objBooking.Save(ref rsHeader,
                                                        ref rsSegment,
                                                        ref rsPassenger,
                                                        ref rsRemark,
                                                        ref rsPayment,
                                                        ref rsMapping,
                                                        ref rsService,
                                                        ref rsTax,
                                                        ref rsFees,
                                                        ref createTickets,
                                                        ref readBooking,
                                                        ref readOnly,
                                                        ref bSetLock,
                                                        ref bCheckSeatAssignment,
                                                        ref bCheckSessionTimeOut);


                    if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                    {
                        result = true;
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                    {
                        throw new BookingSaveException("BOOKINGINUSE");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                    {
                        throw new BookingSaveException("BOOKINGREADERROR");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                    {
                        throw new BookingSaveException("DATABASEACCESS");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                    {
                        throw new BookingSaveException("DUPLICATESEAT", "277");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                    {
                        throw new BookingSaveException("SESSIONTIMEOUT", "H005");
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }
            return result;
        }

        public bool UpdateContactDetail(string strBookingId,
                                    Entity.Booking.BookingHeader header,
                                    string strUserId,
                                    string agencyCode,
                                    bool bNoVat)
        {
            tikAeroProcess.Booking objBooking = null;
            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;

            string strXml = string.Empty;
            bool createTickets = false;
            bool readBooking = false;
            bool readOnly = false;
            bool bSetLock = false;
            bool bCheckSeatAssignment = false;
            bool bCheckSessionTimeOut = false;
            string paymentType = string.Empty;
            string language = "EN";
            string currency = string.Empty;
            entity.Booking bookingResponse = new entity.Booking();
            entity.Booking booking = new entity.Booking();
            var objSegmentIdMapping = new List<KeyValuePair<string, string>>();
            bool result = false;

            tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                }

                if (!String.IsNullOrEmpty(strBookingId))
                {
                    //read booking first for prepare RS 
                    if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                        string.Empty,
                        0,
                        ref rsHeader,
                        ref rsSegment,
                        ref rsPassenger,
                        ref rsRemark,
                        ref rsPayment,
                        ref rsMapping,
                        ref rsService,
                        ref rsTax,
                        ref rsFees,
                        ref rsQuote,
                        false,
                        false,
                        "{" + strUserId.ToUpper() + "}",
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        string.Empty,
                        string.Empty) == true)
                    {
                        // convert rs to booking for preparing header and mapping  value
                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);
                            currency = booking.Header.CurrencyRcd;
                            language = booking.Header.LanguageRcd;
                        }
                    }


                        if (header.BookingId == RecordsetHelper.ToGuid(rsHeader, "booking_id"))
                        {
                            rsHeader.Fields["contact_name"].Value = header.ContactName;
                            rsHeader.Fields["contact_email"].Value = header.ContactEmail;
                            rsHeader.Fields["phone_mobile"].Value = header.PhoneMobile;

                            rsHeader.Fields["phone_home"].Value = header.PhoneHome;
                            rsHeader.Fields["phone_business"].Value = header.PhoneBusiness;
                            rsHeader.Fields["phone_fax"].Value = header.PhoneFax;
                            rsHeader.Fields["mobile_email"].Value = header.MobileEmail;

                            rsHeader.Fields["language_rcd"].Value = header.LanguageRcd;
                            rsHeader.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                            
                        }


                    // save booking
                    enumSaveError = objBooking.Save(ref rsHeader,
                                                        ref rsSegment,
                                                        ref rsPassenger,
                                                        ref rsRemark,
                                                        ref rsPayment,
                                                        ref rsMapping,
                                                        ref rsService,
                                                        ref rsTax,
                                                        ref rsFees,
                                                        ref createTickets,
                                                        ref readBooking,
                                                        ref readOnly,
                                                        ref bSetLock,
                                                        ref bCheckSeatAssignment,
                                                        ref bCheckSessionTimeOut);


                    if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                    {
                        result = true;
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                    {
                        throw new BookingSaveException("BOOKINGINUSE");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                    {
                        throw new BookingSaveException("BOOKINGREADERROR");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                    {
                        throw new BookingSaveException("DATABASEACCESS");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                    {
                        throw new BookingSaveException("DUPLICATESEAT", "277");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                    {
                        throw new BookingSaveException("SESSIONTIMEOUT", "H005");
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }
            return result;
        }



        public bool UpdateTicket(string strBookingId,
                                    IList<Entity.Booking.Mapping> mapping,
                                    string strUserId,
                                    string agencyCode,
                                    bool bNoVat)
        {
            tikAeroProcess.Booking objBooking = null;
            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;

            string strXml = string.Empty;
            bool createTickets = false;
            bool readBooking = false;
            bool readOnly = false;
            bool bSetLock = false;
            bool bCheckSeatAssignment = false;
            bool bCheckSessionTimeOut = false;
            string paymentType = string.Empty;
            string language = "EN";
            string currency = string.Empty;
            entity.Booking bookingResponse = new entity.Booking();
            entity.Booking booking = new entity.Booking();
            var objSegmentIdMapping = new List<KeyValuePair<string, string>>();

            tikAeroProcess.BookingSaveError enumSaveError = tikAeroProcess.BookingSaveError.OK;

            bool result = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                }

                if (!String.IsNullOrEmpty(strBookingId))
                {

                    //read booking first for prepare RS 
                    if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                        string.Empty,
                        0,
                        ref rsHeader,
                        ref rsSegment,
                        ref rsPassenger,
                        ref rsRemark,
                        ref rsPayment,
                        ref rsMapping,
                        ref rsService,
                        ref rsTax,
                        ref rsFees,
                        ref rsQuote,
                        false,
                        false,
                        "{" + strUserId.ToUpper() + "}",
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        string.Empty,
                        string.Empty) == true)
                    {
                        // convert rs to booking for preparing header and mapping  value
                        if (rsHeader != null && rsHeader.RecordCount > 0)
                        {
                            booking.Header = booking.Header.FillBooking(rsHeader);
                            currency = booking.Header.CurrencyRcd;
                            language = booking.Header.LanguageRcd;
                        }
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                        {
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        }
                    }


                    foreach (entity.Mapping mm in mapping)
                    {
                        foreach (entity.Mapping m in booking.Mappings)
                        {
                            if (mm.BookingSegmentId == m.BookingSegmentId && mm.PassengerId == m.PassengerId)
                            {
                                m.BaggageWeight = mm.BaggageWeight;
                                m.PieceAllowance = mm.PieceAllowance;
                                m.EndorsementText = mm.EndorsementText;
                                m.RestrictionText = mm.RestrictionText;
                            }
                        }
                    }


                    foreach (entity.Mapping m in booking.Mappings)
                    {
                        rsMapping.MoveFirst();
                        while (!rsMapping.EOF)
                        {
                            if (m.PassengerId == RecordsetHelper.ToGuid(rsMapping, "passenger_id")
                                &&
                                m.BookingSegmentId == RecordsetHelper.ToGuid(rsMapping, "booking_segment_id")
                                )
                            {
                                rsMapping.Fields["endorsement_text"].Value = m.EndorsementText;
                                rsMapping.Fields["restriction_text"].Value = m.RestrictionText;
                                rsMapping.Fields["baggage_weight"].Value = m.BaggageWeight;
                                rsMapping.Fields["piece_allowance"].Value = m.PieceAllowance;
                                rsMapping.Fields["update_by"].Value = new Guid(strUserId).ToRsString();
                                break;
                            }
                            rsMapping.MoveNext();
                        }
                    }


                    // save booking
                    enumSaveError = objBooking.Save(ref rsHeader,
                                                        ref rsSegment,
                                                        ref rsPassenger,
                                                        ref rsRemark,
                                                        ref rsPayment,
                                                        ref rsMapping,
                                                        ref rsService,
                                                        ref rsTax,
                                                        ref rsFees,
                                                        ref createTickets,
                                                        ref readBooking,
                                                        ref readOnly,
                                                        ref bSetLock,
                                                        ref bCheckSeatAssignment,
                                                        ref bCheckSessionTimeOut);


                    if (enumSaveError == tikAeroProcess.BookingSaveError.OK)
                    {
                        result = true;
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGINUSE)
                    {
                        throw new BookingSaveException("BOOKINGINUSE");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.BOOKINGREADERROR)
                    {
                        throw new BookingSaveException("BOOKINGREADERROR");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DATABASEACCESS)
                    {
                        throw new BookingSaveException("DATABASEACCESS");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.DUPLICATESEAT)
                    {
                        throw new BookingSaveException("DUPLICATESEAT", "277");
                    }
                    else if (enumSaveError == tikAeroProcess.BookingSaveError.SESSIONTIMEOUT)
                    {
                        throw new BookingSaveException("SESSIONTIMEOUT", "H005");
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }
            return result;
        }


        public entity.Booking GetQuoteSummary(IList<entity.FlightSegment> flights, IList<entity.Passenger> passengers, string agencyCode, string language, string currencyCode, bool bNovat)
        {
            tikAeroProcess.Fares objFare = null;
            ADODB.Recordset rsFlightSegment = null;
            ADODB.Recordset rsPassengers = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsQuotes = null;
            ADODB.Recordset rsFee = null;

            string strFlightId = string.Empty;
            string strBoardpoint = string.Empty;
            string strOffpoint = string.Empty;
            string strAirline = string.Empty;
            string strFlight = string.Empty;
            string strBoardingClass = string.Empty;
            string strBookingClass = string.Empty;
            DateTime dtFlight = Convert.ToDateTime("12/30/1899");
            string strFareId = string.Empty;
            bool bRefundable = false;
            bool bGroupBooking = false;
            bool bNonRevenue = false;
            string strSegmentId = string.Empty;
            short iIdReduction = 0;
            string strFareType = "FARE";
  
            entity.Booking booking = new entity.Booking();

            try
            {
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFare = (tikAeroProcess.Fares)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objFare = new tikAeroProcess.Fares(); }

                if (flights != null && flights.Count > 0)
                {
                    flights.ToRecordset(ref rsFlightSegment);
                    passengers.ToRecordset(ref rsPassengers);

                    if (objFare.FareQuote(ref rsPassengers,
                                           agencyCode,
                                           ref rsTax,
                                           ref rsMapping,
                                           ref rsQuotes,
                                           ref rsFlightSegment,
                                           ref strFlightId,
                                           ref strBoardpoint,
                                           ref strOffpoint,
                                           ref strAirline,
                                           ref strFlight,
                                           ref strBoardingClass,
                                           ref strBookingClass,
                                           ref dtFlight,
                                           ref strFareId,
                                           ref currencyCode,
                                           ref bRefundable,
                                           ref bGroupBooking,
                                           ref bNonRevenue,
                                           ref strSegmentId,
                                           ref iIdReduction,
                                           ref strFareType,
                                           ref language,
                                           ref bNovat) == true)
                    {

                        if (rsFlightSegment != null && rsFlightSegment.RecordCount > 0)
                            booking.Segments = booking.Segments.FillBooking(rsFlightSegment);
                        if (rsPassengers != null && rsPassengers.RecordCount > 0)
                            booking.Passengers = booking.Passengers.FillBooking(rsPassengers);
                        if (rsTax != null && rsTax.RecordCount > 0)
                            booking.Taxs = booking.Taxs.FillBooking(rsTax);
                        if (rsMapping != null && rsMapping.RecordCount > 0)
                            booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                        if (rsQuotes != null && rsQuotes.RecordCount > 0)
                            booking.Quotes = booking.Quotes.FillBooking(rsQuotes);
                        if (rsFee != null && rsFee.RecordCount > 0)
                            booking.Fees = booking.Fees.FillBooking(rsFee);
                    }
                }

            }
            catch (BookingException bookingExc)
            {
                throw bookingExc;
            }
            finally
            {
                if (objFare != null)
                {
                    Marshal.FinalReleaseComObject(objFare);
                    objFare = null;
                }
                RecordsetHelper.ClearRecordset(ref rsMapping, false);
                RecordsetHelper.ClearRecordset(ref rsFee);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsQuotes);
                RecordsetHelper.ClearRecordset(ref rsFee);
            }
            return booking;
        }

        private bool BookingChangeFlight(string strBookingId,
                                            ref Recordset rsHeader,
                                            ref Recordset rsSegment,
                                            ref Recordset rsPassenger,
                                            ref Recordset rsRemark,
                                            ref Recordset rsPayment,
                                            ref Recordset rsMapping,
                                            ref Recordset rsService,
                                            ref Recordset rsTax,
                                            ref Recordset rsFees,
                                            ref Recordset rsQuote,
                                            IList<entity.Flight> flights,
                                            string strUserId,
                                            string agencyCode,
                                            string  currency,
                                            string  language,
                                            bool bNoVat)
        {
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsFlights = RecordsetHelper.FabricateFlightRecordset();

            entity.Booking booking = new entity.Booking();

           // Boolean resultChange = false;
            Boolean resultAdd = false;

            try
            {
                {
                    if (rsSegment != null && rsSegment.RecordCount > 0)
                    {
                        booking.Segments = booking.Segments.FillBooking(rsSegment);
                    }
                    if (rsMapping != null && rsMapping.RecordCount > 0)
                    {
                        booking.Mappings = booking.Mappings.FillBooking(rsMapping);
                    }

                    if (flights != null && flights.Count > 0)
                    {
                        List<entity.Flight> flightsSort = flights.OrderByDescending(f => f.DepartureDate).ToList();

                        for (int i = 0; i < flightsSort.Count; i++)
                        {
                            rsSegment.Filter = "segment_status_rcd = 'HK' AND od_origin_rcd = '" + flightsSort[i].OdOriginRcd + "' AND od_destination_rcd = '" + flightsSort[i].OdDestinationRcd + "'";

                            if (rsSegment.RecordCount > 0)
                            {
                                rsSegment.Sort = "departure_date DESC";
                                rsSegment.MoveFirst();

                                while (!rsSegment.EOF)
                                {
                                    Guid curValue = RecordsetHelper.ToGuid(rsSegment, "booking_segment_id");
                                    String flightCheckInStatus = RecordsetHelper.ToString(rsSegment, "flight_check_in_status_rcd");

                                    if (booking.IsCheckedPassenger(curValue))
                                    {
                                        throw new ModifyBookingException("Passenger check in status as CHECKED", "L003");
                                    }
                                    if (booking.IsBoardedPassenger(curValue))
                                    {
                                        throw new ModifyBookingException("Passenger check in status as BOARDED", "L004");
                                    }
                                    if (booking.IsFlownPassenger(curValue))
                                    {
                                        throw new ModifyBookingException("Passenger check in status as FLOWN", "L005");
                                    }
                                    if (flightCheckInStatus.Equals("CLOSED"))
                                    {
                                        throw new ModifyBookingException("Flight segment closed", "L002");
                                    }

                                    if (flightsSort.Count(fgt => fgt.ExchangedSegmentId.ToString() == "" + curValue.ToString() + "") == 0)
                                    {
                                        flightsSort[i].ExchangedSegmentId = curValue;
                                        break;
                                    }

                                    rsSegment.MoveNext();
                                }
                            }
                            else
                            {
                                throw new ModifyBookingException("Route not found", "P009");
                            }
                        }

                        flightsSort.ToRecordset(ref rsFlights);
                    }

                    // Call AddChangeFlight for Add new FlightSegment
                    resultAdd = AddChangeFlight(ref rsHeader,
                                            ref rsPassenger,
                                            ref rsSegment,
                                            ref rsMapping,
                                            ref rsPayment,
                                            ref rsTax,
                                            ref rsFees,
                                            ref rsService,
                                            ref rsQuote,
                                            ref rsRemark,
                                            ref rsFlights,
                                            agencyCode,
                                            strBookingId,
                                            language,
                                            currency,
                                            strUserId,
                                            bNoVat);

                    if (resultAdd)
                    {
                        booking.Fees = AddChangeFees(agencyCode, currency, strBookingId, ref rsHeader, ref rsSegment, ref rsPassenger, ref rsFees, ref rsRemark, ref rsMapping, ref rsService, language, bNoVat);
                    }
                    else
                    {
                        throw new ModifyBookingException("CAN NOT ADD FLIGHT", "P008");
                    }

                }
            }
            catch (ModifyBookingException bookingExc)
            {
                throw bookingExc;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
            }
            return resultAdd;
        }

        private Boolean AddChangeFlight(ref Recordset rsHeader,
                                        ref Recordset rsPassenger, 
                                        ref Recordset rsSegment, 
                                        ref Recordset rsMapping,
                                        ref Recordset rsPayment,
                                        ref Recordset rsTax,
                                        ref Recordset rsFees,
                                        ref Recordset rsService, 
                                        ref Recordset rsQuote, 
                                        ref Recordset rsRemark, 
                                        ref Recordset rsFlights,
                                        String agencyCode,
                                        String strBookingId, 
                                        String language, 
                                        String currency, 
                                        String strUserId,
                                        Boolean bNoVat)
        {
            Boolean result = false;
            tikAeroProcess.Booking objBooking = null;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                }

                DateTime dtFlight = DateTime.Parse("1899/12/30");
                Boolean bGroupBooking = false;
                Boolean bWaitlist = false;

                result = objBooking.AddFlight(ref rsPassenger,
                                              ref rsSegment,
                                              ref rsMapping,
                                              ref rsTax,
                                              ref rsService,
                                              agencyCode,
                                              ref rsQuote,
                                              ref rsRemark,
                                              ref rsFlights,
                                              dtFlight,
                                              strBookingId,
                                              "", "", "", "", "", "", "", "",
                                              language,
                                              currency,
                                              false,
                                              false,
                                              bGroupBooking,
                                              bWaitlist,
                                              false,
                                              false,
                                              false,
                                              "", "", "",
                                              0,
                                              strUserId,
                                              false,
                                              false,
                                              "",
                                              bNoVat);

                if (!result)
                    throw new BookingException("Add Change Flight : " + result);

            }
            catch
            {
                throw;
            }

            return result;
        }

        
        
        private IList<Entity.Booking.Fee> AddChangeFees(String agencyCode, 
                                   String currency, 
                                   String strBookingId, 
                                   ref Recordset rsHeader, 
                                   ref Recordset rsSegment, 
                                   ref Recordset rsPassenger, 
                                   ref Recordset rsFees, 
                                   ref Recordset rsRemark, 
                                   ref Recordset rsMapping, 
                                   ref Recordset rsService, 
                                   String language, 
                                   Boolean bNoVat)
        {
            Boolean result = false;
            tikAeroProcess.Fees objFees = null;
            Entity.Booking.Booking booking = new Entity.Booking.Booking();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objFees = new tikAeroProcess.Fees();
                }

                result = objFees.BookingChange(agencyCode, currency, strBookingId, ref rsHeader, ref rsSegment, ref rsPassenger, ref rsFees, ref rsRemark, ref rsMapping, rsService, language, bNoVat);

                if (rsFees != null && rsFees.RecordCount > 0)
                {
                    booking.Fees = booking.Fees.FillBooking(rsFees);
                }

            }
            catch
            {
                throw;
            }
            return booking.Fees;
        }


        //
        public bool COBPayment(Booking booking,
                               IList<entity.Payment> payments,
                               string strUserId, string agencyCode, string currencyRcd, decimal totalOutStanding)
        {
            tikAeroProcess.Booking objBooking = null;
            tikAeroProcess.clsCreditCard objCreditCard = null;
            tikAeroProcess.Payment objPayment = null;
            tikAeroProcess.Fees objFees = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            ADODB.Recordset rsAllocation = null;
            string paymentType = string.Empty;
            string strXml = string.Empty;
            string strBookingId = booking.Header.BookingId.ToString();
            bool bNoVat = Convert.ToBoolean(booking.Header.NoVatFlag == 0 ? false : true);
            bool IsPaymentSave = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.clsCreditCard", _server);
                    objCreditCard = (tikAeroProcess.clsCreditCard)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;

                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                    objCreditCard = new tikAeroProcess.clsCreditCard();
                    objFees = new tikAeroProcess.Fees();
                    objPayment = new tikAeroProcess.Payment();
                }

                    //prepare obj booking

                    if (totalOutStanding > 0)
                    {
                        // get empty rs object
                        rsPayment = objPayment.GetEmpty();

                        if (payments != null && payments.Count == 1) // Fix =1 cause phase 1 suport only single payment
                        {
                            payments[0].DocumentAmount = totalOutStanding;
                            payments[0].PaymentAmount = totalOutStanding;
                            payments[0].ReceivePaymentAmount = totalOutStanding;
                            payments[0].BookingId = new Guid(strBookingId);
                            payments[0].RecordLocator = booking.Header.RecordLocator;

                            if (string.IsNullOrEmpty(payments[0].CurrencyRcd))
                                payments[0].CurrencyRcd = currencyRcd;

                            if (string.IsNullOrEmpty(payments[0].ReceiveCurrencyRcd))
                                payments[0].ReceiveCurrencyRcd = currencyRcd;

                            // agency from booking  to GWS
                            payments[0].AgencyCode = agencyCode;
                            payments[0].CreateBy = new Guid(strUserId);
                            payments[0].UpdateBy = new Guid(strUserId);

                            payments.ToRecordset(ref rsPayment);
                            booking.Payments = booking.Payments.FillBooking(rsPayment);
                            paymentType = booking.Payments[0].FormOfPaymentRcd;
                            rsAllocation = RecordsetHelper.FabricatePaymentAllocationRecordset();
                            IList<Entity.PaymentAllocation> paymentAllocations = null;
                            paymentAllocations = booking.CreateAllocation();
                            PaymentService ps = new PaymentService(_server, _user, _pass, _domain);

                            if (paymentAllocations != null)
                            {
                                //Check Payment Type
                                if (paymentType.Trim().ToUpper().Equals("VOUCHER"))
                                {
                                    IsPaymentSave = ps.PaymentVoucher(totalOutStanding, booking.Mappings, booking.Fees, payments, paymentAllocations, agencyCode,strUserId);
                                }
                                else if (paymentType.Trim().ToUpper().Equals("CRAGT") || paymentType.Trim().ToUpper().Equals("INV"))
                                {
                                    IsPaymentSave = ps.PaymentAgencyAccount(totalOutStanding,paymentType, agencyCode, strUserId, payments, paymentAllocations);
                                }
                                else if (paymentType.Trim().ToUpper().Equals("CC"))
                                {
                                    IsPaymentSave = ps.PaymentCreditCard(booking.Header, payments, paymentAllocations);
                                }
                            }
                        }
                        else
                        {
                            if (payments == null || payments.Count == 0)
                                throw new ModifyBookingException("Required payment information", "P015");
                            else if (payments.Count > 1)
                                throw new ModifyBookingException("Do not support multiple payment", "P016");
                        }
                    }
                    else if (totalOutStanding == 0)
                    {
                        IsPaymentSave = true;
                    }
                    else if (totalOutStanding < 0)
                    {
                        throw new ModifyBookingException("Outstanding balance less than 0", "P017");
                    }
            }
            catch
            {
                throw new ModifyBookingException("Payment fail", "P001"); ;
               // throw;
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                if (objCreditCard != null)
                {
                    Marshal.FinalReleaseComObject(objCreditCard);
                    objCreditCard = null;
                }
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }
                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsQuote);
            }
            return IsPaymentSave;
        }

        // segment id mapping NOT suport connection flight, Oct 2015
        public List<KeyValuePair<string, string>> FindNewSegmentFlightChange(IList<Entity.Booking.FlightSegment> fs)
        {
            var listSegmentIdMapping = new List<KeyValuePair<string, string>>();
            var listOldSegId = new List<KeyValuePair<string, string>>();
            var listNewSegId = new List<KeyValuePair<string, string>>();

            if (fs != null && fs.Count > 0)
            {
                for (int i = 0; i < fs.Count; i++)
                {
                    if (fs[i].SegmentStatusRcd.ToUpper().Equals("NN"))
                    {
                        listNewSegId.Add(new KeyValuePair<string, string>(fs[i].BookingSegmentId.ToString(), fs[i].OdOriginRcd + fs[i].OdDestinationRcd));
                    }
                    else if (fs[i].SegmentStatusRcd.ToUpper().Equals("XK"))
                    {
                        listOldSegId.Add(new KeyValuePair<string, string>(fs[i].BookingSegmentId.ToString(), fs[i].OdOriginRcd + fs[i].OdDestinationRcd));
                    }

                    // if found status US return null
                    else if (fs[i].SegmentStatusRcd.ToUpper().Equals("US"))
                    {
                        listSegmentIdMapping = null;
                        return listSegmentIdMapping;
                    }
                }

                //for (int i = 0; i < fs.Count; i++)
                //{
                //    if (fs[i].SegmentStatusRcd.ToUpper().Equals("XK"))
                //    {
                //        listOldSegId.Add(new KeyValuePair<string, string>(fs[i].BookingSegmentId.ToString(), fs[i].OdOriginRcd + fs[i].OdDestinationRcd));
                //    }
                //}

                // if found status US return null
                //for (int i = 0; i < fs.Count; i++)
                //{
                //    if (fs[i].SegmentStatusRcd.ToUpper().Equals("US"))
                //    {
                //        listSegmentIdMapping = null;
                //        return listSegmentIdMapping;
                //    }
                //}

                foreach (var elementNew in listNewSegId)
                {
                    foreach (var elementOld in listOldSegId)
                    {
                        if (elementNew.Value.Equals(elementOld.Value))
                        {
                            listSegmentIdMapping.Add(new KeyValuePair<string, string>(elementOld.Key, elementNew.Key));
                        }
                    }
                }

            }

            return listSegmentIdMapping;
        }
        
        public bool SetRSnewSegmentId2(List<KeyValuePair<string, string>> objSegmentIdMapping, ref Recordset rs)
        {
            bool result = false;
            if (objSegmentIdMapping != null && objSegmentIdMapping.Count > 0)
            {
                foreach (var element in objSegmentIdMapping)
                {
                    if (rs != null && rs.RecordCount > 0)
                    {
                        rs.MoveFirst();
                        while (!rs.EOF)
                        {
                            if (RecordsetHelper.ToGuid(rs, "booking_segment_id") == new Guid(element.Key))
                            {
                                rs.Fields["booking_segment_id"].Value = "{" + element.Value + "}";
                                result = true;
                            }
                            rs.MoveNext();
                        }
                    }
                }
            }
            return result;
        }

        private bool PaymentVoucher(Booking booking,
                            IList<Entity.Booking.Payment> payments,
                            IList<Entity.PaymentAllocation> paymentAllocations,
                            decimal totalOutStanding)
        {
            tikAeroProcess.Payment objPayment = null;
            ADODB.Recordset rsPayment = null;

            bool result = false;
            ADODB.Recordset rsAllocation = null;
            ADODB.Recordset rsRefundVoucher = null;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objPayment = new tikAeroProcess.Payment();
                }

                // get empty rs object
                rsPayment = objPayment.GetEmpty();

                payments.ToRecordset(ref rsPayment);
                booking.Payments = booking.Payments.FillBooking(rsPayment);

                rsAllocation = RecordsetHelper.FabricatePaymentAllocationRecordset();

                Entity.Voucher voucher = new Entity.Voucher();
                VoucherService vcService = new VoucherService(_server, _user, _pass, _domain);
                bool IsValidVoucher = true;

                if (payments != null && payments.Count > 0)
                {
                    foreach (Entity.Booking.Payment p in payments)
                    {
                        voucher.VoucherNumber = p.DocumentNumber;
                        voucher.VoucherPassword = p.DocumentPassword;
                        voucher.CurrencyRcd = p.CurrencyRcd;
                      //  voucher = vcService.GetVoucher(voucher,null,null);

                        if (voucher == null)
                        {
                            IsValidVoucher = false;
                            throw new ModifyBookingException("INVALID VOUCHER", "P011");
                        }
                        else
                        {
                            rsPayment.MoveFirst();
                            while (!rsPayment.EOF)
                            {
                                rsPayment.Fields["voucher_payment_id"].Value = Guid.Parse(voucher.VoucherId).ToRsString();
                                rsPayment.MoveNext();
                            }
                        }
                    }
                }

                //
                if (IsValidVoucher)
                {
                    paymentAllocations.ToRecordset(ref rsAllocation);

                    if (voucher.ValidateVoucherEnough(totalOutStanding) == true)
                    {
                        result = objPayment.Save(ref rsPayment, ref rsAllocation, ref rsRefundVoucher);

                        if (result)
                        {
                            // createTickets = true;
                        }
                        else
                        {
                            throw new ModifyBookingException("PAYMENT FAIL", "P001");
                        }
                    }
                    else
                    {
                        throw new ModifyBookingException("VALUE NOT ENOUGH", "P002");
                    }
                }
                else
                {
                    throw new ModifyBookingException("PAYMENT DOCUMENT NOT FOUND", "P003");
                }
            }
            catch
            {
                throw new ModifyBookingException("PAYMENT FAIL", "P001");
            }
            finally
            {
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }

                RecordsetHelper.ClearRecordset(ref rsPayment);
            }

            return result;
        }


        private bool InfantOverLimit(int numberOfInfant, string flightId, string originRcd, string destinationRcd, string boardingClass)
        {
            tikAeroProcess.Inventory objInventory = null;
            ADODB.Recordset rsInfantLimit = null;
            bool result = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Inventory", _server);
                    objInventory = (tikAeroProcess.Inventory)Activator.CreateInstance(remote);
                    remote = null;

                }
                else
                {
                    objInventory = new tikAeroProcess.Inventory();
                }

                rsInfantLimit = objInventory.GetFlightLegInfants(ref flightId,
                                                                    ref originRcd,
                                                                    ref destinationRcd,
                                                                    ref boardingClass);


                if (rsInfantLimit != null && rsInfantLimit.RecordCount > 0)
                {
                    // if (rsInfantLimit.Fields["infant_capacity"].Value != System.DBNull.Value)
                    {
                        if (Convert.ToInt32(rsInfantLimit.Fields["infant_capacity"].Value) > 0 && numberOfInfant > 0)
                        {
                            if (Convert.ToInt32(rsInfantLimit.Fields["available_infant"].Value) < numberOfInfant)
                            {
                                result = true;
                            }
                            else
                            {
                                result = false;
                            }

                        }
                    }
                }

            }
            catch
            {

            }
            finally
            {
                if (objInventory != null)
                {
                    Marshal.FinalReleaseComObject(objInventory);
                    objInventory = null;
                }
                RecordsetHelper.ClearRecordset(ref rsInfantLimit);

            }
            return result;
        }

        private Entity.Booking.Booking RsFillBooking(ADODB.Recordset rsHeader,
                                                ADODB.Recordset rsSegment,
                                                ADODB.Recordset rsPassenger,
                                                ADODB.Recordset rsFees,
                                                ADODB.Recordset rsRemark,
                                                ADODB.Recordset rsQuote,
                                                ADODB.Recordset rsTax,
                                                ADODB.Recordset rsService,
                                                ADODB.Recordset rsMapping)
        {
            Entity.Booking.Booking booking = new Booking();

            if (rsHeader != null && rsHeader.RecordCount > 0)
            {
                booking.Header = booking.Header.FillBooking(rsHeader);
            }
            if (rsSegment != null && rsSegment.RecordCount > 0)
            {
                booking.Segments = booking.Segments.FillBooking(rsSegment);
            }
            if (rsPassenger != null && rsPassenger.RecordCount > 0)
            {
                booking.Passengers = booking.Passengers.FillBooking(rsPassenger);
            }
            if (rsFees != null && rsFees.RecordCount > 0)
            {
                booking.Fees = booking.Fees.FillBooking(rsFees);
            }
            if (rsRemark != null && rsRemark.RecordCount > 0)
            {
                booking.Remarks = booking.Remarks.FillBooking(rsRemark);
            }
            if (rsQuote != null && rsQuote.RecordCount > 0)
            {
                booking.Quotes = booking.Quotes.FillBooking(rsQuote);
            }
            if (rsTax != null && rsTax.RecordCount > 0)
            {
                booking.Taxs = booking.Taxs.FillBooking(rsTax);
            }
            if (rsService != null && rsService.RecordCount > 0)
            {
                booking.Services = booking.Services.FillBooking(rsService);
            }
            if (rsMapping != null && rsMapping.RecordCount > 0)
            {
                booking.Mappings = booking.Mappings.FillBooking(rsMapping);
            }

            return booking;
        }

    }
}
