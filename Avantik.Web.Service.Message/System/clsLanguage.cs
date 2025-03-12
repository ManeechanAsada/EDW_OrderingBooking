using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace Avantik.Web.Service.Message
{
    public class Language 
    {
        #region Property

        public string LanguageRcd { get; set; }
        public string DisplayName { get; set; }
        public string CharacterSet { get; set; }

        #endregion
    }
}
