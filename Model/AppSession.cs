using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace techlink_workspace.Model
{
    public static class AppSession
    {
        public static UserModel CurrentUser { get; set; }

        public static bool IsLoggedIn => CurrentUser != null;

        public static void Clear() => CurrentUser = null;
    }
}
