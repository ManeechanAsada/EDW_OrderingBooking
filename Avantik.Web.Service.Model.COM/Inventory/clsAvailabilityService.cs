using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using Avantik.Web.Service.Model.Contract;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.COMHelper;

namespace Avantik.Web.Service.Model.COM.Flight
{
    public class AvailabilityService : RunComplus, IAvailabilityService
    {
        string _server = string.Empty;
        public AvailabilityService(string server, string user, string pass, string domain)
            : base(user, pass, domain)
        {
            _server = server;
        }
        public Entity.Flight.Availabilities GetAvailability(string otherPassengerType, 
                                                            string boardingClass, 
                                                            string bookingClass, 
                                                            string dayTimeIndicator,
                                                            string returnDayTimeIndicator,
                                                            string agencyCode, 
                                                            string currencyCode, 
                                                            string transitPoint, 
                                                            string promotionCode, 
                                                            AvailabilityFareTypes fareType, 
                                                            string languageCode, 
                                                            string ipAddress, 
                                                            string originRcd, 
                                                            string destinationRcd, 
                                                            string odOriginRcd, 
                                                            string odDestinationRcd, 
                                                            DateTime fromDate, 
                                                            DateTime toDate, 
                                                            DateTime returnFromDate, 
                                                            DateTime returnToDate, 
                                                            DateTime bookingDate, 
                                                            decimal maxAmount, 
                                                            byte adult, 
                                                            byte child, 
                                                            byte infant, 
                                                            byte other, 
                                                            Guid flightId, 
                                                            Guid fareId, 
                                                            short nonStopOnly, 
                                                            short includeDeparted, 
                                                            short includeCancelled, 
                                                            short includeWaitlisted, 
                                                            short includeSoldOut, 
                                                            short includeFares, 
                                                            short refundable,
                                                            short returnRefundable,
                                                            short groupFares, 
                                                            short iTFaresOnly, 
                                                            short staffFares, 
                                                            bool applyFareLogic, 
                                                            bool showClose, 
                                                            short unknownTransit, 
                                                            bool mapWithFares, 
                                                            bool noVat, 
                                                            string transactionReference)
        {
            Entity.Flight.Availabilities availabilities = null;
            IList<Entity.Flight.Availability> availability = null;

            tikAeroProcess.Inventory objInventory = null;
            ADODB.Recordset rsOutward = null;
            ADODB.Recordset rsReturn = null;
            ADODB.Recordset rsLogic = null;

            try
            {
                //Initialize declaration for all fare.
                bool bLowestFare = true;
                bool bLowestClass = true;
                bool bLowestGroup = true;
                bool bSort = false;
                bool bDelete = true;
                bool bClosed = false;
                bool bNonStopOnly = Convert.ToBoolean(nonStopOnly);
                bool bIncludeDeparted = Convert.ToBoolean(includeDeparted);
                bool bIncludeCancelled = Convert.ToBoolean(includeCancelled);
                bool bIncludeWaitlisted = Convert.ToBoolean(includeWaitlisted);
                bool bIncludeSoldOut = Convert.ToBoolean(includeSoldOut);
                bool bRefundable = Convert.ToBoolean(refundable);
                bool bReturnRefundable = Convert.ToBoolean(returnRefundable);
                bool bGroupFares = Convert.ToBoolean(groupFares);
                bool bITFaresOnly = Convert.ToBoolean(iTFaresOnly);
                bool bStaffFares = Convert.ToBoolean(staffFares);
                bool bUnknownTransit = Convert.ToBoolean(unknownTransit);
                bool bSkipFarelogic = false;
                bool bReturn = false;

                short iAdult = adult;
                short iChild = child;
                short iInfant = infant;
                short iOther = other;

                double dblMaxAmount = Convert.ToDouble(maxAmount);

                string strFlightId = string.Empty;
                string strFareId = string.Empty;
                string strFareType = "FARE";

                //Check if it is search by date range. If yes skip fare logic.
                if (fromDate != toDate && returnFromDate != returnToDate)
                {
                    applyFareLogic = false;
                }

                //Set fare type parameter.
                if (fareType == AvailabilityFareTypes.LOWEST)
                {
                    bLowestFare = true;
                    bLowestClass = false;
                    bLowestGroup = true;
                    bSort = false;
                    bDelete = true;
                    bClosed = false;
                }
                else if (fareType == AvailabilityFareTypes.LOWESTCLASS)
                {
                    bLowestFare = true;
                    bLowestClass = true;
                    bLowestGroup = true;
                    bSort = false;
                    bDelete = true;
                    bClosed = false;
                }
                else if (fareType == AvailabilityFareTypes.LOWESTGROUP)
                {
                    bLowestFare = false;
                    bLowestClass = false;
                    bLowestGroup = true;
                    bSort = false;
                    bDelete = true;
                    bClosed = false;
                }
                else if (fareType == AvailabilityFareTypes.REDEMPTION)
                {
                    bLowestFare = false;
                    bLowestClass = false;
                    bLowestGroup = false;
                    bSort = false;
                    bDelete = false;
                    bClosed = false;
                    strFareType = "POINT";
                }

                //Call availability COM plus.
                
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Inventory", _server);
                    objInventory = (tikAeroProcess.Inventory)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objInventory = new tikAeroProcess.Inventory(); }

                if (objInventory != null)
                {

                    //Assign value from parameter.
                    if (flightId.Equals(Guid.Empty) == false)
                    {
                        strFlightId = flightId.ToString();
                    }

                    if (fareId.Equals(Guid.Empty) == false)
                    {
                        strFareId = fareId.ToString();
                    }

                    if (returnFromDate.Equals(DateTime.MinValue) == false && returnToDate.Equals(DateTime.MinValue) == false)
                    {
                        bReturn = true;
                    }

                    rsOutward = objInventory.GetAvailability(originRcd,
                                                             destinationRcd,
                                                             fromDate,
                                                             toDate,
                                                             bookingDate,
                                                             ref iAdult,
                                                             ref iChild,
                                                             ref iInfant,
                                                             ref iOther,
                                                             ref otherPassengerType,
                                                             ref boardingClass,
                                                             ref bookingClass,
                                                             ref dayTimeIndicator,
                                                             ref agencyCode,
                                                             ref currencyCode,
                                                             ref strFlightId,
                                                             ref strFareId,
                                                             ref dblMaxAmount,
                                                             ref bNonStopOnly,
                                                             ref bIncludeDeparted,
                                                             ref bIncludeCancelled,
                                                             ref bIncludeWaitlisted,
                                                             ref bIncludeSoldOut,
                                                             ref bRefundable,
                                                             ref bGroupFares,
                                                             ref bITFaresOnly,
                                                             ref bStaffFares,
                                                             ref applyFareLogic,
                                                             ref bUnknownTransit,
                                                             ref transitPoint,
                                                             ref returnFromDate,
                                                             ref returnToDate,
                                                             ref rsReturn,
                                                             ref mapWithFares,
                                                             ref bReturnRefundable,
                                                             ref returnDayTimeIndicator,
                                                             ref promotionCode,
                                                             ref strFareType,
                                                             ref languageCode,
                                                             ref ipAddress,
                                                             ref noVat);


                    if (applyFareLogic == true)
                    {
                        bSkipFarelogic = false;
                    }
                    else
                    {
                        bSkipFarelogic = true;
                    }
                    objInventory.FilterAvailability(ref rsOutward,
                                                    iAdult,
                                                    iChild,
                                                    iInfant,
                                                    iOther,
                                                    ref rsReturn,
                                                    ref rsLogic,
                                                    originRcd,
                                                    destinationRcd,
                                                    fromDate,
                                                    returnFromDate,
                                                    bookingDate,
                                                    bReturn,
                                                    bLowestFare,
                                                    bLowestClass,
                                                    bLowestGroup,
                                                    bClosed,
                                                    bSort,
                                                    bDelete,
                                                    bSkipFarelogic,
                                                    bReturnRefundable);

                    //Fill Data To Availability object.
                    if (rsOutward != null && rsOutward.RecordCount > 0)
                    {
                        availability = new List<Entity.Flight.Availability>();
                        availabilities = new Entity.Flight.Availabilities();
                        

                        //Fill outbound information.
                        availability.FillAvailability(rsOutward, "OUTBOUND");
                        availabilities.FlightAvailabilityOutbound = availability;
                        availability = null;

                        //Fill return information.
                        if (rsReturn != null && rsReturn.RecordCount > 0)
                        {
                            availability = new List<Entity.Flight.Availability>();
                            
                            availability.FillAvailability(rsReturn, "RETURN");
                            availabilities.FlightAvailabilityReturn = availability;
                            
                            availability = null;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                //Release com object.
                Marshal.FinalReleaseComObject(objInventory);
                objInventory = null;

                //Clear recordset resource.
                RecordsetHelper.ClearRecordset(ref rsOutward);
                RecordsetHelper.ClearRecordset(ref rsReturn);
                RecordsetHelper.ClearRecordset(ref rsLogic);
            }

            return availabilities;
        }
    }
}
