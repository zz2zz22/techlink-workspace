using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace techlink_workspace.Model
{
    public class UserModel
    {
        public string User_id { get; set; }
        public string User_code { get; set; }
        public string User_fullName { get; set; }
        public string User_password { get; set; }
        public int User_type { get; set; }
        public int User_status { get; set; }
    }
}
