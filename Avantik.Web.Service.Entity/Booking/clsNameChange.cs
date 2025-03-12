using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Entity

{
    public class NameChange
    {
        public Guid PassengerId { get; set; }
        public string TitleRcd { get; set; }
        public string Lastname { get; set; }
        public string Firstname { get; set; }
        public string Middlename { get; set; }
        public string GenderTypeRcd { get; set; }
        public DateTime DateOfBirth { get; set; }
    }
}
