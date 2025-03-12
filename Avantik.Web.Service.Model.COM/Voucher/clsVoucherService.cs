using Avantik.Web.Service.Model.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Agency;
using Avantik.Web.Service.COMHelper;
using System.Runtime.InteropServices;
using Avantik.Web.Service.Model.COM.Extension;
using Avantik.Web.Service.Exception.Booking;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Helpers;


namespace Avantik.Web.Service.Model.COM
{
    public class VoucherService : RunComplus, IVoucherService
    {
        string _server = string.Empty;
        public VoucherService(string server, string user, string pass, string domain)
            : base(user, pass, domain)
        {
            _server = server;
        }

        public Voucher GetVoucher(Voucher vc,IList<Entity.Booking.Mapping> Mappings, IList<Entity.Booking.Fee> Fees,string agencyCode, string userId)
        {
            tikAeroProcess.Miscellaneous objMiscellaneous = null;
            tikAeroProcess.Payment objPayment = null;
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsVoucher = null;
            Voucher voucher = null;

            DateTime dtFrom = DateTime.MinValue;
            DateTime dtTo = DateTime.MinValue;

            ADODB.Recordset rsMappings = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFlight = null;
            ADODB.Recordset rsQuote = null;
            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strCurrency = string.Empty;

            short iAdult = 0;
            short iChild = 0;
            short iInfant = 0;
            short iOther = 0;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Miscellaneous", _server);
                    objMiscellaneous = (tikAeroProcess.Miscellaneous)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                }
                else
                { 
                    objMiscellaneous = new tikAeroProcess.Miscellaneous();
                    objPayment = new tikAeroProcess.Payment();
                    objBooking = new tikAeroProcess.Booking(); 
                }

                // get empty
                string result = objBooking.GetEmpty(ref rsHeader,
                                   ref rsSegment,
                                   ref rsPassenger,
                                   ref rsRemark,
                                   ref rsPayment,
                                   ref rsMappings,
                                   ref rsService,
                                   ref rsTax,
                                   ref rsFees,
                                   ref rsFlight,
                                   ref rsQuote,
                                   ref strBookingId,
                                   ref agencyCode,
                                   ref strCurrency,
                                   ref iAdult,
                                   ref iChild,
                                   ref iInfant,
                                   ref iOther,
                                   ref strOther,
                                   ref userId);

                if (result != string.Empty)
                {
                    if (Fees != null && Fees.Count > 0)
                    {
                        Fees.ToRecordset(ref rsFees);
                    }
                    if (Mappings != null && Mappings.Count > 0)
                    {
                        Mappings.ToRecordset(ref rsMappings);
                    }
                }

                rsVoucher = objMiscellaneous.GetVouchers("",
                                                  vc.VoucherNumber,
                                                  "",
                                                  "",
                                                  "",
                                                  "",
                                                  "",
                                                  vc.CurrencyRcd,
                                                  vc.VoucherPassword,
                                                  true,
                                                  false,
                                                  false,
                                                  false,
                                                  false,
                                                  false,
                                                  false,
                                                  ref dtFrom,
                                                  ref dtTo,
                                                  "");

              // validate 
             ADODB.Recordset rsVoucherAllocation = null;
             rsVoucherAllocation = objPayment.ValidateVoucher(ref rsVoucher, rsMappings, rsFees, String.Empty);

             //rsVoucherAllocation.MoveFirst();
             //string currency_rcd = RecordsetHelper.ToString(rsVoucherAllocation, "currency_rcd");
             //double max_amount = RecordsetHelper.ToDouble(rsVoucherAllocation, "max_amount");
             //double dif_amount = RecordsetHelper.ToDouble(rsVoucherAllocation, "dif_amount");
             //double outstanding_amount = RecordsetHelper.ToDouble(rsVoucherAllocation, "outstanding_amount");


              if (rsVoucher != null && rsVoucher.RecordCount > 0)
              {
                  voucher = voucher.FillVoucher(rsVoucher);
              }

              string ValidateVoucherFlag = ConfigHelper.ToString("ValidateVoucherFlag");

              if (ValidateVoucherFlag.ToUpper() == "TRUE")
              {
                  if (rsVoucherAllocation == null)
                  {
                      voucher.VoucherStatusRcd = "NP";
                  }
                  else if (rsVoucherAllocation.RecordCount == 0)
                  {
                      voucher.VoucherStatusRcd = "NP";
                  }
                  else
                  {
                      rsVoucherAllocation.Filter = "dif_amount > 0";
                      if (rsVoucherAllocation.RecordCount > 0)
                      {
                          voucher.VoucherStatusRcd = "NP";
                      }
                  }
              }

                // insert VC Session
              string strVoucherSessionId = Guid.NewGuid().ToString();
              if (objMiscellaneous.InsertVoucherSession(strVoucherSessionId, rsVoucher.Fields["voucher_id"].Value.ToString(), agencyCode, string.Empty, userId))
              {
              }
            }
            catch (BookingException ex)
            {
                throw ex;
            }
            finally
            {
                if (objMiscellaneous != null)
                {
                    Marshal.FinalReleaseComObject(objMiscellaneous);
                    objMiscellaneous = null;
                }
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }

                RecordsetHelper.ClearRecordset(ref rsVoucher);
                RecordsetHelper.ClearRecordset(ref rsMappings);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }

            return voucher;
        }
    }
}
