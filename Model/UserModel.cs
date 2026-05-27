namespace techlink_workspace.Model
{
    public class UserModel
    {
        public string User_id { get; set; }
        public string User_code { get; set; }
        public string User_fullName { get; set; }
        public string User_password { get; set; }
        public string User_type { get; set; }   // nvarchar(50) — stores UserType ID
        public int User_status { get; set; }
        public int? User_permissionLevel { get; set; }   // 1=Admin, 2=Manager, 3=Normal
    }
}