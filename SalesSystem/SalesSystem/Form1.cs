using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SalesSystem
{
    public partial class Form1 : Form
    {
        int count,a;
        string temp1,temp2,check;
        DataGridViewRowCollection rows;
        DataGridViewRow selectRow;
        DataGridViewColumnCollection columns;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "SELECT P_name, P_price, P_count, P_size FROM Products;";
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
                        string temp = reader["P_size"].ToString();
                        string temp0 = reader["P_name"].ToString();
                        cb品項1.Items.Add(temp0 + " " + temp);
                        rows = dg商品.Rows;
                        rows.Add(new object[] { reader["P_name"], reader["P_size"], reader["P_price"],
                        reader["P_count"]});
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
                qs1 = "select * from Orders;";
                    //引用SqlCommand物件
                using (SqlCommand command = new SqlCommand(qs1, cn))
                {
                    //使用SqlDataReader讀取SqlCommand物件
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        rows = dg訂單.Rows;
                        rows.Add(new object[] { reader["O_id"], reader["O_date"], reader["O_way"],
                    reader["O_name"],reader["O_stat"]});
                    }
                    reader.Close();
                }
            }
        }
        private void 電話預訂_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void btn商品新增_Click(object sender, EventArgs e)
        {
            string count = tb數量.Text.ToString();
            string name = tb商品名稱.Text.ToString();
            string price = tb單價.Text.ToString();
            string size = tb規格.Text.ToString();
            string cs = "";
            string queryString = "";
            string qs = "select P_name,P_size from Products";
            cs = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            queryString = "insert into Products (P_name,P_price,P_count,P_size)" +
                    " values(N'" + name + "',N'" + price + "',N'" + count + "',N'" + size + "')";
            DialogResult R;
            R = MessageBox.Show("您確認要新增" + name + " "+ size + "?", "新增確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (R == DialogResult.Yes)
            {
                using (SqlConnection cn = new SqlConnection(cs))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs, cn))
                    {
                        if (dg商品.Rows.Count != 0)
                        {
                            SqlDataReader reader = command.ExecuteReader();
                            reader.Read();
                            temp1 = reader["P_name"].ToString();
                            temp2 = reader["P_size"].ToString();
                            reader.Close();
                        }
                        if (temp1 != name && temp2 != size)
                        {
                            using (SqlCommand cmd = new SqlCommand(queryString, cn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            MessageBox.Show("已有相同規格商品");
                        }
                    }



                }
                dg商品.Rows.Clear();
                cb品項1.Items.Clear();
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "SELECT P_name, P_price, P_count, P_size FROM Products;";
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string temp = reader["P_size"].ToString();
                            string temp0 = reader["P_name"].ToString();
                            cb品項1.Items.Add(temp0 + " " + temp);
                            rows = dg商品.Rows;
                            rows.Add(new object[] { reader["P_name"], reader["P_size"], reader["P_price"],
                            reader["P_count"]});
                        }
                        reader.Close();
                    }
                }
                tb商品名稱.Text = "";
                tb單價.Text = "";
                tb數量.Text = "0";
                tb規格.Text = "";
            }
            else
            {
                MessageBox.Show("新增取消");
            }

        }

        private void btn欄位清除_Click(object sender, EventArgs e)
        {
            DialogResult R;
            R = MessageBox.Show("您確認要清除欄位?", "清除確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (R == DialogResult.Yes)
            {
                tb商品名稱.Text = "";
                tb單價.Text = "";
                tb數量.Text = "0";
                tb規格.Text = "";
            }
            else
            {
                
            }
        }

        private void btn數量加_Click(object sender, EventArgs e)
        {
            tb數量.Text = (Convert.ToInt32(tb數量.Text) + 1).ToString();
            btn數量減.Enabled = true;
        }

        private void btn數量減_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(tb數量.Text) > 0)
            {
                tb數量.Text = (Convert.ToInt32(tb數量.Text) - 1).ToString();
            }
            else
            {
                btn數量減.Enabled = false;
            }
        }

        private void btn商品修改_Click(object sender, EventArgs e)
        {
            string count = tb數量.Text.ToString();
            string name = tb商品名稱.Text.ToString();
            string price = tb單價.Text.ToString();
            string size = tb規格.Text.ToString();
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "UPDATE Products SET P_count = '" + count + "',P_name = '" + name + "',P_price = '" +
                    price + "',P_size = '" + size + "' where P_name = '" + name + "'";
            DialogResult R;
            R = MessageBox.Show("您確認要修改產品?", "修改確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (R == DialogResult.Yes)
            {
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        command.ExecuteNonQuery();
                    }
                }
                dg商品.Rows.Clear();
                cb品項1.Items.Clear();
                string cs2 = "";
                string qs2 = "";
                cs2 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs2 = "SELECT P_name, P_price, P_count, P_size FROM Products;";
                using (SqlConnection cn = new SqlConnection(cs2))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs2, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string temp = reader["P_size"].ToString();
                            string temp0 = reader["P_name"].ToString();
                            cb品項1.Items.Add(temp0 + " " + temp);
                            rows = dg商品.Rows;
                            rows.Add(new object[] { reader["P_name"], reader["P_size"], reader["P_price"],
                            reader["P_count"]});
                        }
                        reader.Close();
                    }
                }
                MessageBox.Show("修改完成");
            } else
            {
                MessageBox.Show("儲存取消");
            }
        }

        private void btn商品刪除_Click(object sender, EventArgs e)
        {
            string count = tb數量.Text.ToString();
            string name = tb商品名稱.Text.ToString();
            string price = tb單價.Text.ToString();
            string size = tb規格.Text.ToString();
            string cs = "";
            string queryString = "";
            cs = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            queryString = "delete from Products where P_name = '" + name +"' and P_size = '" + size + "'";
            DialogResult R;
            R = MessageBox.Show("您確認要刪除" + name + " " + size + " 這項資料嗎?", "刪除確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (R == DialogResult.Yes)
            {
                using (SqlConnection cn = new SqlConnection(cs))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand(queryString, cn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                dg商品.Rows.Clear();
                cb品項1.Items.Clear();
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "SELECT P_name, P_price, P_count,P_size FROM Products;";
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string temp = reader["P_size"].ToString();
                            string temp0 = reader["P_name"].ToString();                           
                            cb品項1.Items.Add(temp0 + " " + temp);
                            rows = dg商品.Rows;
                            rows.Add(new object[] { reader["P_name"], reader["P_size"], reader["P_price"],
                            reader["P_count"]});
                        }
                        reader.Close();
                    }
                }
                tb商品名稱.Text = "";
                tb單價.Text = "";
                tb數量.Text = "0";
                tb規格.Text = "";
            }
            else
            {
                MessageBox.Show("刪除取消");
            }

        }

        private void btn品項1數量加_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(tb品項1數量.Text) >= Convert.ToInt32(lb庫存.Text))
            {
                MessageBox.Show("目前庫存不足，無法再增加數量了");
            }
            else
            {
                tb品項1數量.Text = (Convert.ToInt32(tb品項1數量.Text) + 1).ToString();
                btn品項1數量減.Enabled = true;
                lb品項1小計.Text = (Convert.ToInt32(tb品項1數量.Text) *
                    Convert.ToInt32(lb品項1單價.Text)).ToString();
            }
        }

        private void btn品項1數量減_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(tb品項1數量.Text) > 0)
            {
                tb品項1數量.Text = (Convert.ToInt32(tb品項1數量.Text) - 1).ToString();
            }
            else
            {
                btn品項1數量減.Enabled = false;
            }
            lb品項1小計.Text = (Convert.ToInt32(tb品項1數量.Text) *
                Convert.ToInt32(lb品項1單價.Text)).ToString();
        }

    
        private void cb品項1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string qs1 = "", cs1 = "";
            if (cb品項1.SelectedIndex != -1)
            check = cb品項1.SelectedItem.ToString();
            char space = ' ';
            string[] str1 = check.Split(space);
            qs1 = "SELECT P_price,P_count FROM Products WHERE P_name = '" + str1[0] + "';";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
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
                        string temp1 = reader[0].ToString();
                        string temp2 = reader[1].ToString();
                        lb品項1單價.Text = temp1;
                        lb庫存.Text = temp2;
                    }
                    reader.Close();
                }
            }
        }


        private void btn賣出入帳_Click(object sender, EventArgs e)
        {
            DialogResult R;
            R = MessageBox.Show("您確認要結帳了嗎?", "賣出確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (R == DialogResult.Yes)
            {
                if (dg門市暫存明細.RowCount != 0)
                {
                    string check = lb訂單號.Text;
                    string cs = "";
                    string queryString = "";
                    cs = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                    queryString = "insert into Orders values (" + check + ", getdate(), '門市取貨', '-', '-', '-', '已處理')";
                    using (SqlConnection cn = new SqlConnection(cs))
                    {
                        cn.Open();
                        using (SqlCommand cmd = new SqlCommand(queryString, cn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    for (int i = 0; i < dg門市暫存明細.RowCount; i++)
                    {
                        char space = ' ';
                        string[] str1 = dg門市暫存明細[0, i].Value.ToString().Split(space);
                        cs = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                        queryString = "insert into OrderDetails values (" + check + ", '" + str1[0] + "', " +
                            dg門市暫存明細[2, i].Value + ", " + dg門市暫存明細[1, i].Value + ")";
                        using (SqlConnection cn = new SqlConnection(cs))
                        {
                            cn.Open();
                            using (SqlCommand cmd = new SqlCommand(queryString, cn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                        queryString = "update Products set P_count = P_count - " + dg門市暫存明細[1, i].Value +
                            " where P_name = '" + str1[0] + "' and P_size = '" + str1[1] + "'";
                        using (SqlConnection cn = new SqlConnection(cs))
                        {
                            cn.Open();
                            using (SqlCommand cmd = new SqlCommand(queryString, cn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("明細中並無商品");
                }
                MessageBox.Show("此筆交易已入帳");
                lb訂單號.Text = (Convert.ToInt32(lb訂單號.Text) + 1).ToString();
                dg門市暫存明細.Rows.Clear();
                lb總計.Text = "0";
                dg商品.Rows.Clear();
            }
            else
            {

            }
            string cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            string qs1 = "SELECT P_name, P_price, P_count, P_size FROM Products;";
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
                        rows = dg商品.Rows;
                        rows.Add(new object[] { reader["P_name"], reader["P_size"], reader["P_price"],
                        reader["P_count"]});
                    }
                    reader.Close();
                }
            }
            qs1 = "select * from Orders;";
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
                        rows = dg訂單.Rows;
                        rows.Add(new object[] { reader["O_id"], reader["O_date"], reader["O_way"],
                        reader["O_name"],reader["O_stat"]});
                    }
                    reader.Close();
                }
            }
        }
        private void btn加入明細_Click(object sender, EventArgs e)
        {
            if (int.TryParse(tb品項1數量.Text,out a) == true)
            {
                rows = dg門市暫存明細.Rows;
                rows.Add(new object[] { cb品項1.SelectedItem.ToString(), tb品項1數量.Text, lb品項1單價.Text,
                        (Convert.ToInt32(lb品項1單價.Text) * Convert.ToInt32(tb品項1數量.Text)).ToString()});
                for (int i = 0; i < dg門市暫存明細.RowCount; i++)
                {
                    count += Convert.ToInt32(dg門市暫存明細[3, i].Value);
                }               
                cb品項1.SelectedIndex = -1;
                tb品項1數量.Text = "0";
                lb品項1小計.Text = "0";
                lb品項1單價.Text = "0";
                lb庫存.Text = "0";
                lb總計.Text = count.ToString();
                lb庫存.Text = (Convert.ToInt32(lb庫存.Text) -
                    Convert.ToInt32(tb品項1數量.Text)).ToString();
            }
            else
            {
                MessageBox.Show("數量不正確無法加入明細");
                cb品項1.SelectedIndex = -1;
                lb品項1單價.Text = "0";
                lb品項1小計.Text = "0";
                lb總計.Text = "0";
                tb品項1數量.Text = "0";
            }

        }

        private void dg商品_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int indexrow = e.RowIndex;
            if(indexrow >= 0)
            {
                selectRow = dg商品.Rows[indexrow];
                tb商品名稱.Text = selectRow.Cells[0].Value.ToString();
                tb規格.Text = selectRow.Cells[1].Value.ToString();
                tb單價.Text = selectRow.Cells[2].Value.ToString();
                tb數量.Text = selectRow.Cells[3].Value.ToString();
            }
        }

        private void btn明細項刪除_Click(object sender, EventArgs e)
        {
            if (lbox銷售明細.SelectedIndex != -1)
            {
                char[] separ = { ' ', ' ' };
                string[] str3 = lbox銷售明細.SelectedItem.ToString().Split(separ);
                string word = "";
                str3[3] = str3[3].Replace("小計", word);
                str3[3] = str3[3].Replace("元", word);
                count -= Convert.ToInt32(str3[3]);
                lb總計.Text = count.ToString();
                lbox銷售明細.Items.RemoveAt(lbox銷售明細.SelectedIndex);
            }else
            {
                MessageBox.Show("請先選取要刪除的明細項");
            }
        }

        private void btn取消交易_Click(object sender, EventArgs e)
        {
            cb品項1.SelectedIndex = -1;
            lb品項1單價.Text = "0";
            lb品項1小計.Text = "0";
            lb總計.Text = "0";
            tb品項1數量.Text = "0";
            lb庫存.Text = "0";
            lbox銷售明細.Items.Clear();
        }
    }
}
