using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Flight;
using Avantik.Web.Service.Infrastructure;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Entity.FormOfPayment;

namespace Avantik.Web.Service.Model.Contract
{
    public interface ISystemModelService
    {
        IList<Country> GetCountry(string language);
        IList<Language> GetLanguage(string language);
        IList<Title> GetTitle(string language);
        IList<Currency> GetCurrency(string language);
        IList<SpecialService> GetSpecialService(string language);
        FormOfPayment GetFormOfPaymentSubType(string type,string language);
        List<Document> GetDocumentType(string language);
    }
}
