using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Data;

namespace techlink_workspace.Repositories
{
    public class SqlSoft
    {
        public SqlConnection conn = DatabaseUtils.GetDBConnection();
        public string sqlExecuteScalarStringHoldConnect(string sql)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (SqlException ex)
            {
                CTMessageBox.Show("Không thể kết nối server SQL Soft!\r\n无法连接到 SQL Soft 服务器！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            String outstring;
            if (conn.State == ConnectionState.Open)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                try
                {
                    outstring = cmd.ExecuteScalar().ToString();
                    return outstring;
                }
                catch (Exception)
                {
                    return String.Empty;
                }
            }
            else
            {
                return String.Empty;
            }
        }
        public string sqlExecuteScalarString(string sql)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (SqlException ex)
            {
                CTMessageBox.Show("Không thể kết nối server SQL Soft!\r\n无法连接到 SQL Soft 服务器！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            String outstring;
            if (conn.State == ConnectionState.Open)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                try
                {
                    outstring = cmd.ExecuteScalar().ToString();
                    conn.Close();
                    return outstring;
                }
                catch (Exception)
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                    return String.Empty;
                }
            }
            else
            {
                return String.Empty;
            }
        }
        public void getComboBoxData(string sql, ref ComboBox cmb)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (SqlException ex)
            {
                CTMessageBox.Show("Không thể kết nối server SQL Soft!\r\n无法连接到 SQL Soft 服务器！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (conn.State == ConnectionState.Open)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter adapter = new SqlDataAdapter();
                try
                {
                    adapter.SelectCommand = cmd;
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    adapter.Dispose();
                    cmd.Dispose();
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        cmb.Items.Add(row[0].ToString());
                    }
                    conn.Close();
                }
                catch (Exception ex)
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                    CTMessageBox.Show("Lỗi khi tải dữ liệu vào combobox!\r\n将数据加载到组合框中时出错！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public void sqlDataAdapterFillDatatableHoldConnect(string sql, ref DataTable dt)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (SqlException ex)
            {
                CTMessageBox.Show("Không thể kết nối server SQL Soft!\r\n无法连接到 SQL Soft 服务器！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (conn.State == ConnectionState.Open)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                try
                {
                    adapter.Fill(dt);
                }
                catch (Exception ex)
                {
                    CTMessageBox.Show("Lỗi khi tải dữ liệu datatable!\r\n加载数据表数据时出错！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public void sqlDataAdapterFillDatatable(string sql, ref DataTable dt)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (SqlException ex)
            {
                CTMessageBox.Show("Không thể kết nối server SQL Soft!\r\n无法连接到 SQL Soft 服务器！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (conn.State == ConnectionState.Open)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                try
                {
                    adapter.Fill(dt);
                }
                catch (Exception ex)
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                    CTMessageBox.Show("Lỗi khi tải dữ liệu datatable!\r\n加载数据表数据时出错！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public void sqlExecuteNonQuery(string sql, string successMessage, string errorMessage)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (SqlException ex)
            {
                CTMessageBox.Show("Không thể kết nối server SQL Soft!\r\n无法连接到 SQL Soft 服务器！\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (conn.State == ConnectionState.Open)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                try
                {
                    cmd.CommandTimeout = 5;
                    cmd.ExecuteNonQuery();
                    ////if (!string.IsNullOrEmpty(successMessage))
                        //Alert(successMessage, Form_Alert.enmType.Success);
                    conn.Close();
                }
                catch (Exception ex)
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                    if (!string.IsNullOrEmpty(errorMessage))
                        CTMessageBox.Show(errorMessage + "\r\n\r\n" + ex.Message, "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
