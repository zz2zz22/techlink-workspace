using System;

namespace techlink_workspace.Model
{
    public class PICModel
    {
        public string PIC_Code { get; set; }
        public string PIC_Name { get; set; }
        /// <summary>Semicolon-separated customer codes, e.g. "KH001;KH002;KH005"</summary>
        public string PIC_CustomerName { get; set; }
        public DateTime? createdate { get; set; }
        public string createby { get; set; }
        public DateTime? updatedate { get; set; }
        public string updateby { get; set; }
    }
}