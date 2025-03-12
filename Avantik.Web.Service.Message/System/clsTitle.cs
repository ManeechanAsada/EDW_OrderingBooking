using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;
namespace Avantik.Web.Service.Message
{
    public class Title 
    {
        #region Property

        public string TitleRcd { get; set; }
        public string DisplayName { get; set; }
        public string GenderTypeRcd { get; set; }

        #endregion
    }
}
