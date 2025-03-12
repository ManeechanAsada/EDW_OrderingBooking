using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Repository;
using Avantik.Web.Service.Repository.Contract.Fares;
using Avantik.Web.Service.Entity.Fares;
using Avantik.Web.Service.Helpers;

namespace Avantik.Web.Service.Model.Fares
{
    public class FareService
    {
        IFareRepository _fareRepository;
        public FareService()
        {
            _fareRepository = Repository.Factory.Fares.FareFactory.CreateInstance();
        }
        public FareService(IFareRepository fareRepository)
        {
            _fareRepository = fareRepository;
        }
        public string GetFareLogicClasses(string originRcd,
                                          string destinationRcd,
                                          DateTime dateOutbound,
                                          DateTime dateReturn,
                                          DateTime bookingDate)
        {
            try
            {
                if (_fareRepository != null)
                {
                    if (string.IsNullOrEmpty(originRcd))
                    {
                        throw new ArgumentException("originRcd required");
                    }
                    else if (string.IsNullOrEmpty(destinationRcd))
                    {
                        throw new ArgumentException("destinationRcd required");
                    }
                    else if (dateOutbound.Equals(DateTime.MinValue))
                    {
                        throw new ArgumentException("dateOutbound required");
                    }
                    else
                    {
                        long lBookDay = 0;
                        long lFlightDay = 0;

                        bool bOneway = false;

                        //Get number of day compare to booking date.
                        lBookDay = Date.DateDiffDay(bookingDate, dateOutbound);
                        if (dateReturn.Equals(DateTime.MinValue) == false)
                        {
                            lFlightDay = Date.DateDiffDay(dateOutbound, dateReturn);
                            bOneway = false;
                        }
                        else
                        {
                            bOneway = true;
                        }

                        FareLogic fareLogic = _fareRepository.GetFareLogicBookingClass(originRcd,
                                                                                    destinationRcd,
                                                                                    dateOutbound,
                                                                                    dateReturn,
                                                                                    bookingDate);
                        if (fareLogic != null)
                        {
                            if (fareLogic.level_1_prior_days < lBookDay)
                            {
                                if (bOneway == true)
                                {
                                    return fareLogic.level_1_oneway_classes;
                                }
                                else
                                {
                                    return fareLogic.level_1_return_classes;
                                }
                            }
                            else if (fareLogic.level_2_prior_days < lBookDay)
                            {
                                if (bOneway == true)
                                {
                                    return fareLogic.level_2_oneway_classes;
                                }
                                else if (fareLogic.level_2_1_days < lFlightDay)
                                {
                                    return fareLogic.level_2_1_return_classes;
                                }
                                else if (fareLogic.level_2_2_days < lFlightDay)
                                {
                                    return fareLogic.level_2_2_shortstay_classes;
                                }
                                else
                                {
                                    return fareLogic.level_2_2_return_classes;
                                }
                            }
                            else
                            {
                                if (bOneway == true)
                                {
                                    return fareLogic.level_3_oneway_classes;
                                }
                                else if (fareLogic.level_3_1_days < lFlightDay)
                                {
                                    return fareLogic.level_3_1_return_classes;
                                }
                                else if (fareLogic.level_3_2_days < lFlightDay)
                                {
                                    return fareLogic.level_3_2_shortstay_classes;
                                }
                                else
                                {
                                    return fareLogic.level_3_2_return_classes;
                                }
                            }
                        }
                    }
                    
                }
                else
                {
                    throw new ArgumentNullException("Fare repository is required.");
                }

                return string.Empty;
            }
            catch
            {
                throw;
            }
        }
    }
}
