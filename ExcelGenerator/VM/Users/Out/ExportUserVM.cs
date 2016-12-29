using ExcelGenerator.VM.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.VM.Out
{
    public class ExportUserVM
    {
        public ExportUserVM()
        {
            Users = new List<UserModelVM>();
        }

        [DisplayName("Usuários")]
        public List<UserModelVM> Users { get; set; }
    }
}
