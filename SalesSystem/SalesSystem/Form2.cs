using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace SalesSystem
{
    public partial class Form2 : Form
    {
        int count, a;
        string temp1, check;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "SELECT P_name, P_price, P_count FROM Products;";
            string qs2 = "SELECT MAX(O_id) FROM OrderDetails";
            using (SqlConnection cn = new SqlConnection(cs1))
            {
                //開啟資料庫
                cn.Open();
                //引用SqlCommand物件
                using (SqlCommand command = new SqlCommand(qs1, cn))
                {
                    //使用SqlDataReader讀取SqlCommand物件
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        string[] temp = { reader[0].ToString(),
                            reader[1].ToString(), reader[2].ToString() };
                        cb電話品項.Items.Add(temp[0]);
                    }
                    reader.Close();
                }
                using (SqlCommand command = new SqlCommand(qs2, cn))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        string temp = reader[0].ToString();
                        if (temp != "")
                        {
                            lb訂單號.Text = (Convert.ToInt32(temp) + 1).ToString();
                        }
                    }
                    reader.Close();
                }
            }
        }
    }
}
