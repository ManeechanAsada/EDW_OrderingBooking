using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Avantik.Web.Service.Helpers;
using Avantik.Web.Service.Helpers.Database;
using Avantik.Web.Service.Repository.Contract.Fares;
using Avantik.Web.Service.Entity.Fares;

namespace Avantik.Web.Service.Repository.Fares
{
    public class FareRepository : IFareRepository
    {
        string _connectionString;
        public FareRepository() { 
        
        }
        public FareRepository(string connectionString)
        {
            _connectionString = connectionString;
        }
        public FareLogic GetFareLogicBookingClass(string originRcd,
                                                               string destinationRcd,
                                                               DateTime dateOutbound,
                                                               DateTime dateReturn,
                                                               DateTime bookingDate)
        {
            try
            {
                FareLogic fareLogic = null;
                SqlDataReader rw;

                //Open DB connection.
                using (DBHelper db = new DBHelper(_connectionString))
                {
                    db.Connect();
                    //Call store procedure.
                    rw = db.ExecDataReaderProc("get_fare_logic_classes",
                                                                       "@origin", originRcd,
                                                                       "@destination", destinationRcd,
                                                                       "@dateOutbound", dateOutbound.GetDateString());

                    if (rw != null && rw.HasRows == true)
                    {
                        //Fill fare login data to farelogic object.
                        fareLogic = new FareLogic();
                        while (rw.Read())
                        {
                            fareLogic.level_1_prior_days = rw.DBToInt32("level_1_prior_days");
                            fareLogic.level_2_prior_days = rw.DBToInt32("level_2_prior_days");
                            fareLogic.level_2_1_days = rw.DBToInt32("level_2_1_days");
                            fareLogic.level_2_2_days = rw.DBToInt32("level_2_2_days");
                            fareLogic.level_3_1_days = rw.DBToInt32("level_3_1_days");
                            fareLogic.level_3_2_days = rw.DBToInt32("level_3_2_days");

                            fareLogic.level_1_oneway_classes = rw.DBToString("level_1_oneway_classes");
                            fareLogic.level_1_return_classes = rw.DBToString("level_1_return_classes");
                            fareLogic.level_2_oneway_classes = rw.DBToString("level_2_oneway_classes");
                            fareLogic.level_2_1_return_classes = rw.DBToString("level_2_1_return_classes");
                            fareLogic.level_2_2_shortstay_classes = rw.DBToString("level_2_2_shortstay_classes");
                            fareLogic.level_2_2_return_classes = rw.DBToString("level_2_2_return_classes");
                            fareLogic.level_3_oneway_classes = rw.DBToString("level_3_oneway_classes");
                            fareLogic.level_3_1_return_classes = rw.DBToString("level_3_1_return_classes");
                            fareLogic.level_3_2_shortstay_classes = rw.DBToString("level_3_2_shortstay_classes");
                            fareLogic.level_3_2_return_classes = rw.DBToString("level_3_2_return_classes");
                            break;
                        }

                    }
                }

                return fareLogic;
            }
            catch
            {
                throw;
            }

        }
    }
}
