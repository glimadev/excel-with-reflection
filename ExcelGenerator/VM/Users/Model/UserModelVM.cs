using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.VM.Model
{
    public class UserModelVM
    {
        [DisplayName("Name")]
        public string FullName { get; set; }

        [DisplayName("Registration Date")]
        public DateTime CreatedDate { get; set; }

        [DisplayName("City")]
        public string City { get; set; }

        [DisplayName("State")]
        public string State { get; set; }

        //Hidden from excel generator reflection
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string UserId { get; set; }
    }
}
