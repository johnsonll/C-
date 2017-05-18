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
        int count, a, j,indexrows, orderrows, memberrows, icount;
        string temp1,temp2,check;
        DataGridViewRowCollection rows;
        DataGridViewRow selectRow;
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
            string qs2 = "SELECT MAX(O_id) FROM Orders";
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
                qs2 = "SELECT MAX(Cus_id) FROM Customers";
                using (SqlCommand command = new SqlCommand(qs2, cn))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        string temp = reader[0].ToString();
                        if (temp != "")
                        {
                            lb會員號.Text = (Convert.ToInt32(temp) + 1).ToString();
                        }
                    }
                    reader.Close();
                }
                qs1 = "select * from Orders where O_way != '來店';";
                    //引用SqlCommand物件
                using (SqlCommand command = new SqlCommand(qs1, cn))
                {
                    //使用SqlDataReader讀取SqlCommand物件
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        rows = dg訂單.Rows;
                        rows.Add(new object[] { reader["O_id"], reader["O_date"], reader["O_way"],
                    reader["O_name"],reader["O_tel"],reader["O_address"],reader["O_stat"]});
                    }
                    reader.Close();
                }              
                qs1 = "select * from Customers";
                using (SqlCommand command = new SqlCommand(qs1, cn))
                {
                    //使用SqlDataReader讀取SqlCommand物件
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        rows = dg會員資料表1.Rows;
                        rows.Add(new object[] { reader["Cus_id"], reader["Cus_name"], reader["Cus_tel"], reader["Cus_address"], reader["Cus_relationship"] });
                        rows = dg會員資料表2.Rows;
                        rows.Add(new object[] { reader["Cus_name"], reader["Cus_tel"], reader["Cus_address"] });
                    }
                    reader.Close();
                }
            }
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
            j = 0;
            string qs1 = "", cs1 = "";
            if (cb品項1.SelectedIndex != -1)
            {
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
                            if (dg門市暫存明細.RowCount > 0)
                            {
                                if (rbtn來店消費.Checked || rbtn電話預訂門市取貨.Checked)
                                {
                                    for (int i = 0; i < dg門市暫存明細.RowCount; i++)
                                    {
                                        if (cb品項1.SelectedItem.ToString().Equals(dg門市暫存明細.Rows[i].Cells[0].Value.ToString()))
                                        {
                                            temp2 = (Convert.ToInt32(temp2) -
                                                Convert.ToInt32(dg門市暫存明細.Rows[i].Cells[1].Value)).ToString();
                                        }
                                    }
                                    lb庫存.Text = temp2;
                                }else
                                {
                                    lb庫存.Text = temp2;
                                }
                            }
                            else
                            {
                                lb庫存.Text = temp2;
                            }
                        }
                        reader.Close();
                    }
                }
            }
        }

        private void btn加入常客_Click(object sender, EventArgs e)
        {
            string cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            string queryString = "insert into Customers values ('" + tb銷售顧客稱謂.Text + "','"
                + tb銷售顧客電話.Text + "','" + tb銷售顧客地址.Text + "','-')";
            string qs = "select * from Customers";
            try
            {
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    //開啟資料庫
                    cn.Open();
                    //引用SqlCommand物件
                    using (SqlCommand cmd = new SqlCommand(queryString, cn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                dg會員資料表2.Rows.Clear();
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs, cn))
                    {
                        //使用SqlDataReader讀取SqlCommand物件
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            rows = dg會員資料表2.Rows;
                            rows.Add(new object[] { reader["Cus_name"], reader["Cus_tel"], reader["Cus_address"] });
                        }
                        reader.Close();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("已有相同會員");
            }
            tb銷售顧客稱謂.Text = "";
            tb銷售顧客地址.Text = "";
            tb銷售顧客電話.Text = "";
        }

        private void dg會員資料表2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int indexrow = e.RowIndex;
            if (indexrow >= 0)
            {
                selectRow = dg會員資料表2.Rows[indexrow];
                tb銷售顧客稱謂.Text = selectRow.Cells[0].Value.ToString();
                tb銷售顧客電話.Text = selectRow.Cells[1].Value.ToString();
                tb銷售顧客地址.Text = selectRow.Cells[2].Value.ToString();
            }
        }

        private void btn模糊搜尋_Click(object sender, EventArgs e)
        {
            dg會員資料表2.Rows.Clear();
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "SELECT * FROM Customers where Cus_name like '%" + tb銷售顧客稱謂.Text + "%'";
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
                        rows = dg會員資料表2.Rows;
                        rows.Add(new object[] { reader["Cus_name"], reader["Cus_tel"], reader["Cus_address"] });
                    }
                    reader.Close();
                }
            }
        }

        private void btn重置搜尋_Click(object sender, EventArgs e)
        {
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "SELECT * FROM Customers";
            dg會員資料表2.Rows.Clear();
            using (SqlConnection cn = new SqlConnection(cs1))
            {
                cn.Open();
                using (SqlCommand command = new SqlCommand(qs1, cn))
                {
                    //使用SqlDataReader讀取SqlCommand物件
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        rows = dg會員資料表2.Rows;
                        rows.Add(new object[] { reader["Cus_name"], reader["Cus_tel"], reader["Cus_address"] });
                    }
                    reader.Close();
                }
            }
        }

        private void btn訂單搜尋_Click(object sender, EventArgs e)
        {
            dg訂單.Rows.Clear();
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            if(rbtn未處理.Checked && ckbox來店.Checked && ckbox宅配.Checked)
                qs1 = "SELECT * FROM Orders where O_way like '%來店' and O_way != '來店' or O_way like '%宅%' and O_stat = '未處理' or O_stat like '已%未%';";
            else if (rbtn未處理.Checked && ckbox來店.Checked)
                qs1 = "SELECT * FROM Orders where O_way like '%來店' and O_way != '來店' and O_stat = '未處理' or O_stat like '已%未%';";
            else if (rbtn未處理.Checked && ckbox宅配.Checked)
                qs1 = "SELECT * FROM Orders where O_way like '宅配' and O_stat = '未處理' or O_stat like '已%未%';";
            else if (rbtn未處理.Checked)
                qs1 = "SELECT * FROM Orders where O_stat = '未處理' or O_stat like '已%未%';";
            if(rbtn已處理.Checked && ckbox來店.Checked && ckbox宅配.Checked)
                qs1 = "SELECT * FROM Orders where O_way like '%來店' or O_way like '%宅%' and O_stat like '已%' and O_stat not like '已%未%' and O_way != '來店';";
            else if (rbtn已處理.Checked && ckbox來店.Checked)
                qs1 = "SELECT * FROM Orders where O_way like '%來店' and O_stat like '已%' and O_stat not like '已%未%' and O_way != '來店';";
            else if (rbtn已處理.Checked && ckbox宅配.Checked)
                qs1 = "SELECT * FROM Orders where O_way like '宅配' and O_stat like '已%' and O_stat not like '已%未%';";
            else if (rbtn已處理.Checked)
                qs1 = "SELECT * FROM Orders where O_stat like '已%';";
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
                        reader["O_name"], reader["O_tel"], reader["O_address"], reader["O_stat"]});
                    }
                    reader.Close();
                }
            }
        }

        private void btn清除搜尋結果_Click(object sender, EventArgs e)
        {
            dg訂單.Rows.Clear();
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "select * from Orders;";
            using (SqlConnection cn = new SqlConnection(cs1))
            {
                cn.Open();
                using (SqlCommand command = new SqlCommand(qs1, cn))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        rows = dg訂單.Rows;
                        rows.Add(new object[] { reader["O_id"], reader["O_date"], reader["O_way"],
                    reader["O_name"],reader["O_tel"],reader["O_address"],reader["O_stat"]});
                    }
                    reader.Close();
                }
            }
            ckbox來店.Checked = false;
            ckbox宅配.Checked = false;
            rbtn已處理.Checked = true;
            rbtn未處理.Checked = false;
        }

        private void btn已取貨_Click(object sender, EventArgs e)
        {
            if (dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已來店取貨" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已宅配發貨" || 
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已取消")
            {
                MessageBox.Show("此訂單狀態已無法更改");
            }
            else
            {
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "UPDATE Orders SET O_stat = '已來店取貨' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                DialogResult R;
                R = MessageBox.Show("您確認要把此訂單狀態改為已取貨嗎?", "修改確認",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (R == DialogResult.Yes)
                {
                    dg訂單.Rows[orderrows].Cells[6].Value = "已來店取貨";
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
        private void btn已發貨_Click(object sender, EventArgs e)
        {
            if (dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已來店取貨" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已宅配並收款" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已取消")
            {
                MessageBox.Show("此訂單狀態已無法更改");
            }
            else if(dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已收款未宅配")
            {
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "UPDATE Orders SET O_stat = '已宅配並收款' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                DialogResult R;
                R = MessageBox.Show("您確認此訂單已宅配了嗎?", "修改確認",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (R == DialogResult.Yes)
                {
                    dg訂單.Rows[orderrows].Cells[6].Value = "已宅配並收款";
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
            else if (dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "未處理")
            {
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "UPDATE Orders SET O_stat = '已宅配未收款' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                DialogResult R;
                R = MessageBox.Show("您確認此訂單已宅配了嗎?", "修改確認",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (R == DialogResult.Yes)
                {
                    dg訂單.Rows[orderrows].Cells[6].Value = "已宅配未收款";
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
        private void btn訂單取消_Click(object sender, EventArgs e)
        {
            if (dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "未處理")
            {
                string cs1 = "";
                string qs1 = "";
                List<string> name = new List<string>();
                List<string> qty = new List<string>();
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "UPDATE Orders SET O_stat = '已取消' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                DialogResult R;
                R = MessageBox.Show("您確認要把此訂單取消?", "取消確認",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (R == DialogResult.Yes)
                {
                    dg訂單.Rows[orderrows].Cells[6].Value = "已取消";
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            command.ExecuteNonQuery();
                        }
                    }
                    qs1 = "select * from OrderDetails where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            //使用SqlDataReader讀取SqlCommand物件
                            SqlDataReader reader = command.ExecuteReader();
                            while (reader.Read())
                            {
                                name.Add(reader["P_name"].ToString());
                                qty.Add(reader["OD_qty"].ToString());
                            }
                            reader.Close();
                        }
                        if (dg訂單.Rows[orderrows].Cells[2].Value.ToString() == "預約來店")
                        {
                            for (int i = 0; i < name.Count; i++)
                            {
                                qs1 = "update Products set P_count = P_count + " + qty[i] + " where P_name = '" + name[i] + "'";
                                using (SqlCommand command = new SqlCommand(qs1, cn))
                                    command.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }else
            {
                MessageBox.Show("此訂單已為入帳或發貨後的訂單，故無法取消");
            }
        }

        private void btn姓名搜尋_Click(object sender, EventArgs e)
        {
            dg會員資料表1.Rows.Clear();
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "SELECT * FROM Customers where Cus_name like '%" + tb會員姓名.Text + "%'";
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
                        rows = dg會員資料表1.Rows;
                        rows.Add(new object[] { reader["Cus_name"], reader["Cus_tel"], reader["Cus_address"], reader["Cus_relationship"] });
                    }
                    reader.Close();
                }
            }
        }

        private void dg會員資料表1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            memberrows = e.RowIndex;
            if (memberrows >= 0)
            {
                selectRow = dg會員資料表1.Rows[memberrows];
                tb會員姓名.Text = selectRow.Cells[1].Value.ToString();
                tb會員電話.Text = selectRow.Cells[2].Value.ToString();
                tb會員地址.Text = selectRow.Cells[3].Value.ToString();
                tb會員關係.Text = selectRow.Cells[4].Value.ToString();
            }
        }

        private void btn新增會員_Click(object sender, EventArgs e)
        {
            string name = tb會員姓名.Text.ToString();
            string address = tb會員地址.Text.ToString();
            string tel = tb會員電話.Text.ToString();
            string rel = tb會員關係.Text.ToString();
            string cs = "";
            string queryString = "";
            string qs = "select Cus_name, Cus_tel from Customers";
            if (tb會員關係.Text == "")
                rel = "-";
            cs = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            queryString = "insert into Customers (Cus_id,Cus_name,Cus_tel,Cus_address,Cus_relationship)" +
                    " values("+ lb會員號.Text + ",N'" + name + "',N'" + tel + "',N'" + address + "',N'" + rel + "')";
            DialogResult R;
            R = MessageBox.Show("您確認要新增" + name + "?", "新增確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (R == DialogResult.Yes)
            {
                using (SqlConnection cn = new SqlConnection(cs))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs, cn))
                    {
                        if (dg會員資料表1.Rows.Count != 0)
                        {
                            SqlDataReader reader = command.ExecuteReader();
                            reader.Read();
                            temp1 = reader[0].ToString();
                            temp2 = reader[1].ToString();
                            reader.Close();
                        }
                        if (temp1 != name && temp2 != tel)
                        {
                            using (SqlCommand cmd = new SqlCommand(queryString, cn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            MessageBox.Show("已有相同會員");
                        }
                    }
                }
                dg會員資料表1.Rows.Clear();
                dg會員資料表2.Rows.Clear();
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "select * from Customers";
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            rows = dg會員資料表1.Rows;
                            rows.Add(new object[] { reader[0], reader[1], reader[2],
                            reader[3]});
                            rows = dg會員資料表2.Rows;
                            rows.Add(new object[] { reader[1], reader[2], reader[3]});
                        }
                        reader.Close();
                    }
                }
                tb會員姓名.Text = "";
                tb會員地址.Text = "";
                tb會員關係.Text = "";
                tb會員電話.Text = "";
            }
            else
            {
                MessageBox.Show("新增取消");
            }
        }

        private void btn會員修改資料_Click(object sender, EventArgs e)
        {
            string name = tb會員姓名.Text.ToString();
            string address = tb會員地址.Text.ToString();
            string tel = tb會員電話.Text.ToString();
            string rel = tb會員關係.Text.ToString();
            string cs1 = "";
            string qs1 = "";
            cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            qs1 = "UPDATE Customers SET Cus_name = '" + name + "',Cus_tel = '" + tel + "',Cus_address = '" +
                    address + "',Cus_relationship = '" + rel + "' where Cus_id = " + dg會員資料表1.Rows[memberrows].Cells[0].Value;
            DialogResult R;
            R = MessageBox.Show("您確認要修改會員?", "修改確認",
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
                dg會員資料表1.Rows.Clear();
                dg會員資料表2.Rows.Clear();
                string cs2 = "";
                string qs2 = "";
                cs2 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs2 = "SELECT * FROM Customers";
                using (SqlConnection cn = new SqlConnection(cs2))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs2, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            rows = dg會員資料表1.Rows;
                            rows.Add(new object[] { reader[0], reader[1], reader[2],
                            reader[3], reader[4]});
                            rows = dg會員資料表2.Rows;
                            rows.Add(new object[] { reader[1], reader[2], reader[3] });
                        }
                        reader.Close();
                    }
                }
                MessageBox.Show("修改完成");
            }
            else
            {
                MessageBox.Show("儲存取消");
            }
        }

        private void btn刪除會員_Click(object sender, EventArgs e)
        {
            string name = tb會員姓名.Text.ToString();
            string address = tb會員地址.Text.ToString();
            string tel = tb會員電話.Text.ToString();
            string rel = tb會員關係.Text.ToString();
            string cs = "";
            string queryString = "";
            cs = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            queryString = "delete from Customers where Cus_id = "+ dg會員資料表1.Rows[memberrows].Cells[0].Value;
            DialogResult R;
            R = MessageBox.Show("您確認要刪除 " + name + " 這個會員嗎?", "刪除確認",
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
                dg會員資料表1.Rows.Clear();
                dg會員資料表2.Rows.Clear();
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
                            rows = dg會員資料表1.Rows;
                            rows.Add(new object[] { reader[0], reader[1], reader[2],
                            reader[3], reader[4]});
                            rows = dg會員資料表2.Rows;
                            rows.Add(new object[] { reader[1], reader[2], reader[3] });
                        }
                        reader.Close();
                    }
                }
                tb會員地址.Text = "";
                tb會員姓名.Text = "";
                tb會員關係.Text = "";
                tb會員電話.Text = "";
            }
            else
            {
                MessageBox.Show("刪除取消");
            }
        }

        private void btn修改訂單_Click(object sender, EventArgs e)
        {
            string name = tb訂單姓名.Text.ToString();
            string address = tb訂單地址.Text.ToString();
            string tel = tb訂單電話.Text.ToString();
            string cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
            string qs1 = "";
            DialogResult R;
            R = MessageBox.Show("您確認要修改訂單?", "修改確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            try {
                if (R == DialogResult.Yes)
                {
                    if (rbtn訂單來店.Checked)
                        qs1 = "UPDATE Orders SET O_name = '" + name + "',O_tel = '" + tel + "',O_address = '" +
                        address + "', O_way = '預約來店' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                    else if (rbtn訂單宅配.Checked)
                        qs1 = "UPDATE Orders SET O_name = '" + name + "',O_tel = '" + tel + "',O_address = '" +
                        address + "', O_way = '宅配' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            command.ExecuteNonQuery();
                        }
                    }
                    if (rbtn訂單來店.Checked)
                    {
                        dg訂單.Rows[orderrows].Cells[2].Value = rbtn訂單來店.Text;
                    }
                    else if (rbtn訂單宅配.Checked)
                    {
                        dg訂單.Rows[orderrows].Cells[2].Value = rbtn訂單宅配.Text;
                    }
                    dg訂單.Rows[orderrows].Cells[3].Value = tb訂單姓名.Text.ToString();
                    dg訂單.Rows[orderrows].Cells[4].Value = tb訂單電話.Text.ToString();
                    dg訂單.Rows[orderrows].Cells[5].Value = tb訂單地址.Text.ToString();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("資料有誤");
            }
        }

        private void btn先付款_Click(object sender, EventArgs e)
        {
            if (dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已來店取貨" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已宅配並收款" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已取消" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已收款未宅配")
            {
                MessageBox.Show("此訂單狀態已無法更改");
            }
            else if (dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已宅配未收款")
            {
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "UPDATE Orders SET O_stat = '已宅配並收款' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                DialogResult R;
                R = MessageBox.Show("您確認此訂單已付款了嗎?", "修改確認",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (R == DialogResult.Yes)
                {
                    dg訂單.Rows[orderrows].Cells[6].Value = "已宅配並收款";
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
            else if (dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "未處理")
            {
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "UPDATE Orders SET O_stat = '已收款未宅配' where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value;
                DialogResult R;
                R = MessageBox.Show("您確認此訂單已收款了嗎?", "修改確認",
                     MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (R == DialogResult.Yes)
                {
                    dg訂單.Rows[orderrows].Cells[6].Value = "已收款未宅配";
                    using (SqlConnection cn = new SqlConnection(cs1))
                    {
                        cn.Open();
                        using (SqlCommand command = new SqlCommand(qs1, cn))
                        {
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        private void btn查詢銷售紀錄_Click(object sender, EventArgs e)
        {
            dg銷售商品紀錄.Rows.Clear();
            lb總銷售金額.Text = "總銷售金額：";
            dg銷售紀錄.Rows.Clear();
            int total = 0;
            if(tb查詢年.Text == "" && cb查詢月.Text == "")
            {
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "SELECT * FROM Orders where O_stat = '已結帳' or O_stat = '已來店取貨' or O_stat = '已宅配並收款';";
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            rows = dg銷售紀錄.Rows;
                            rows.Add(new object[] {reader["O_id"],reader["O_date"],reader["O_way"]});
                        }
                        reader.Close();
                    }
                }
                qs1 = "SELECT od.OD_totalprice FROM OrderDetails as od inner join Orders as o on o.O_id = od.O_id where o.O_stat = '已結帳' or o.O_stat = '已來店取貨' or o.O_stat = '已宅配並收款';";
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            total += Convert.ToInt32(reader[0]);
                        }
                        reader.Close();
                    }
                }
                lb總銷售金額.Text += total.ToString() + " 元";
                qs1 = "select P_name,sum(OD_qty),sum(OD_totalprice) from OrderDetails where O_id in ( SELECT O_id FROM Orders where O_stat = '已結帳' or O_stat = '已來店取貨' or O_stat = '已宅配並收款') group by P_name";
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            rows = dg銷售商品紀錄.Rows;
                            rows.Add(new object[] { reader[0], reader[1], reader[2] });
                        }
                        reader.Close();
                    }
                }
            }
            else
            {
                if (tb查詢年.Text != "" && cb查詢月.Text == "")
                {
                    try {
                        string cs1 = "";
                        string qs1 = "";
                        cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                        qs1 = "SELECT * FROM Orders where year(O_date) = '"
                            + tb查詢年.Text + "' and (O_stat = '已結帳' or O_stat = '已來店取貨' or O_stat = '已宅配並收款')";
                        using (SqlConnection cn = new SqlConnection(cs1))
                        {
                            cn.Open();
                            using (SqlCommand command = new SqlCommand(qs1, cn))
                            {
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    rows = dg銷售紀錄.Rows;
                                    rows.Add(new object[] { reader["O_id"], reader["O_date"], reader["O_way"] });
                                }
                                reader.Close();
                            }
                        }
                        qs1 = "SELECT od.OD_totalprice FROM OrderDetails as od inner join Orders as o on o.O_id = od.O_id where year(o.O_date) = '"
                            + tb查詢年.Text + "' and (o.O_stat = '已結帳' or o.O_stat = '已來店取貨' or o.O_stat = '已宅配並收款')";
                        using (SqlConnection cn = new SqlConnection(cs1))
                        {
                            cn.Open();
                            using (SqlCommand command = new SqlCommand(qs1, cn))
                            {
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    total += Convert.ToInt32(reader[0]);
                                }
                                reader.Close();
                            }
                            lb總銷售金額.Text += total.ToString() + " 元";
                        }
                        qs1 = "select P_name,sum(OD_qty),sum(OD_totalprice) from OrderDetails where O_id in ( SELECT O_id FROM Orders where year(O_date) = '"
                            + tb查詢年.Text + "' and (O_stat = '已結帳' or O_stat = '已來店取貨' or O_stat = '已宅配並收款')) group by P_name";
                        using (SqlConnection cn = new SqlConnection(cs1))
                        {
                            cn.Open();
                            using (SqlCommand command = new SqlCommand(qs1, cn))
                            {
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    rows = dg銷售商品紀錄.Rows;
                                    rows.Add(new object[] { reader[0], reader[1], reader[2] });
                                }
                                reader.Close();
                            }
                        }

                    }
                    catch
                    {
                        MessageBox.Show("輸入資料有誤");
                    }
                }
                else if (tb查詢年.Text != "" && cb查詢月.Text != "")
                {
                    try {
                        string cs1 = "";
                        string qs1 = "";
                        cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                        qs1 = "SELECT * FROM Orders where year(O_date) = '"
                            + tb查詢年.Text + "' and month(O_date) = '" + cb查詢月.Text + "' and (O_stat = '已結帳' or O_stat = '已來店取貨' or O_stat = '已宅配並收款')";
                        using (SqlConnection cn = new SqlConnection(cs1))
                        {
                            cn.Open();
                            using (SqlCommand command = new SqlCommand(qs1, cn))
                            {
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    rows = dg銷售紀錄.Rows;
                                    rows.Add(new object[] { reader["O_id"], reader["O_date"], reader["O_way"] });
                                }
                                reader.Close();
                            }
                        }
                        qs1 = "SELECT od.OD_totalprice FROM OrderDetails as od inner join Orders as o on o.O_id = od.O_id where year(o.O_date) = '"
                            + tb查詢年.Text + "' and month(o.O_date) = '" + cb查詢月.Text + "' and (o.O_stat = '已結帳' or o.O_stat = '已來店取貨' or o.O_stat = '已宅配並收款')";
                        using (SqlConnection cn = new SqlConnection(cs1))
                        {
                            cn.Open();
                            using (SqlCommand command = new SqlCommand(qs1, cn))
                            {
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    total += Convert.ToInt32(reader[0]);
                                }
                                reader.Close();
                            }
                        }
                        lb總銷售金額.Text += total.ToString() + " 元";
                        qs1 = "select P_name,sum(OD_qty),sum(OD_totalprice) from OrderDetails where O_id in ( SELECT O_id FROM Orders where year(O_date) = '"
                            + tb查詢年.Text + "' and month(O_date) = '" + cb查詢月.Text + "' and (O_stat = '已結帳' or O_stat = '已來店取貨' or O_stat = '已宅配並收款')) group by P_name";
                        using (SqlConnection cn = new SqlConnection(cs1))
                        {
                            cn.Open();
                            using (SqlCommand command = new SqlCommand(qs1, cn))
                            {
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    rows = dg銷售商品紀錄.Rows;
                                    rows.Add(new object[] { reader[0], reader[1], reader[2] });
                                }
                                reader.Close();
                            }
                        }
                    }
                    catch(Exception)
                    {
                        MessageBox.Show("輸入資料有誤");
                    }
                }
                else
                {
                    MessageBox.Show("請正確輸入年份");
                }
            }
        }

        private void btn會員欄位清除_Click(object sender, EventArgs e)
        {
            tb會員姓名.Text = "";
            tb會員地址.Text = "";
            tb會員關係.Text = "";
            tb會員電話.Text = "";
        }

        private void dg訂單_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dg訂單商品.Rows.Clear();
            orderrows = e.RowIndex;
            if(orderrows >= 0)
            {
                selectRow = dg訂單.Rows[orderrows];
                tb訂單姓名.Text = selectRow.Cells[3].Value.ToString();
                tb訂單電話.Text = selectRow.Cells[4].Value.ToString();
                tb訂單地址.Text = selectRow.Cells[5].Value.ToString();
                if (selectRow.Cells[2].Value.ToString() == "預約來店")
                {
                    rbtn訂單來店.Checked = true;
                    rbtn訂單宅配.Checked = false;
                }else if(selectRow.Cells[2].Value.ToString() == "宅配")
                {
                    rbtn訂單來店.Checked = false;
                    rbtn訂單宅配.Checked = true;
                }
                if(selectRow.Cells[6].Value.ToString() == "已結帳" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已來店取貨" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已宅配未收款" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已宅配並收款" ||
                    dg訂單.Rows[orderrows].Cells[6].Value.ToString() == "已取消")
                {
                    tb訂單地址.Enabled = false;
                    tb訂單姓名.Enabled = false;
                    tb訂單電話.Enabled = false;
                    btn修改訂單.Enabled = false;
                    rbtn訂單來店.Enabled = false;
                    rbtn訂單宅配.Enabled = false;
                }else
                {
                    tb訂單地址.Enabled = true;
                    tb訂單姓名.Enabled = true;
                    tb訂單電話.Enabled = true;
                    btn修改訂單.Enabled = true;
                    rbtn訂單來店.Enabled = true;
                    rbtn訂單宅配.Enabled = true;
                }
                string cs1 = "";
                string qs1 = "";
                cs1 = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                qs1 = "SELECT * FROM OrderDetails where O_id = " + dg訂單.Rows[orderrows].Cells[0].Value.ToString();
                using (SqlConnection cn = new SqlConnection(cs1))
                {
                    cn.Open();
                    using (SqlCommand command = new SqlCommand(qs1, cn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            rows = dg訂單商品.Rows;
                            rows.Add(new object[] { reader["P_name"], reader["OD_qty"], reader["P_price"],
                            reader["OD_totalprice"]});
                        }
                        reader.Close();
                    }
                }
            }
        }

        private void dg門市暫存明細_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            indexrows = e.RowIndex;
        }

        private void btn賣出入帳_Click(object sender, EventArgs e)
        {
            DialogResult R;
            R = MessageBox.Show("您確認要成立此筆訂單了嗎?", "訂單確認",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (R == DialogResult.Yes)
            {
                if (dg門市暫存明細.RowCount != 0)
                {
                    string check = lb訂單號.Text;
                    string cs = "";
                    string queryString = "錯誤";
                    cs = "server=localhost\\sqlexpress;database=SalesSystem1;integrated security=SSPI;";
                    if(rbtn來店消費.Checked == false && rbtn來店預約宅配.Checked == false &&
                        rbtn電話預訂宅配.Checked == false && rbtn電話預訂門市取貨.Checked == false)
                    {
                        MessageBox.Show("請先選擇銷售方式");
                    }
                    if (rbtn來店消費.Checked)
                    {
                        if (tb銷售顧客稱謂.Text == "")
                            tb銷售顧客稱謂.Text = "-";
                        if (tb銷售顧客地址.Text == "")
                            tb銷售顧客地址.Text = "-";
                        if (tb銷售顧客電話.Text == "")
                            tb銷售顧客電話.Text = "-";
                        queryString = "insert into Orders values (" + check + ", getdate(), '來店','"
                            + tb銷售顧客稱謂.Text + "','" + tb銷售顧客電話.Text + "','" + tb銷售顧客地址.Text + "', '已結帳')";
                    }
                    if (rbtn來店預約宅配.Checked)
                    {
                        if (tb銷售顧客稱謂.Text == "" || tb銷售顧客電話.Text == "" || tb銷售顧客地址.Text == "")
                        {
                            MessageBox.Show("顧客稱謂或電話地址尚未輸入，無法預約");
                        }
                        else
                        {
                            queryString = "insert into Orders values (" + check + ", getdate(), '宅配','"
                                + tb銷售顧客稱謂.Text + "','" + tb銷售顧客電話.Text + "','" + tb銷售顧客地址.Text + "', '未處理')";
                        }
                    }
                    if (rbtn電話預訂宅配.Checked)
                    {
                        if (tb銷售顧客稱謂.Text == "" || tb銷售顧客電話.Text == "" || tb銷售顧客地址.Text == "")
                        {
                            MessageBox.Show("顧客稱謂或電話地址尚未輸入，無法預約");
                        }
                        else
                        {
                            queryString = "insert into Orders values (" + check + ", getdate(), '宅配','"
                                + tb銷售顧客稱謂.Text + "','" + tb銷售顧客電話.Text + "','" + tb銷售顧客地址.Text + "', '未處理')";
                        }
                    }
                    if (rbtn電話預訂門市取貨.Checked)
                    {
                        if (tb銷售顧客稱謂.Text == "" || tb銷售顧客電話.Text == "")
                        {
                            MessageBox.Show("顧客稱謂或電話尚未輸入，無法預約");
                        }
                        else
                        {
                            queryString = "insert into Orders values (" + check + ", getdate(), '預約來店','"
                                + tb銷售顧客稱謂.Text + "','" + tb銷售顧客電話.Text + "','" + tb銷售顧客地址.Text + "', '未處理')";
                        }
                    }
                    try
                    {
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
                            queryString = "insert into OrderDetails values (" + check + ", '" + str1[0] + " " + str1[1] + "', " +
                                dg門市暫存明細[2, i].Value + ", " + dg門市暫存明細[1, i].Value + ")";
                            using (SqlConnection cn = new SqlConnection(cs))
                            {
                                cn.Open();
                                using (SqlCommand cmd = new SqlCommand(queryString, cn))
                                {
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            if (rbtn來店消費.Checked || rbtn電話預訂門市取貨.Checked)
                            {
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
                        MessageBox.Show("此筆交易已加入訂單");
                        lb訂單號.Text = (Convert.ToInt32(lb訂單號.Text) + 1).ToString();
                        dg門市暫存明細.Rows.Clear();
                        lb總計.Text = "0";
                        dg商品.Rows.Clear();
                        dg訂單.Rows.Clear();
                        tb銷售顧客稱謂.Text = "";
                        tb銷售顧客地址.Text = "";
                        tb銷售顧客電話.Text = "";
                    }
                    catch (Exception)
                    {
                        
                    }
                }
                else
                {
                    MessageBox.Show("明細中並無商品");
                }
                
            }
            else
            {

            }
            dg訂單.Rows.Clear();
            dg商品.Rows.Clear();
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
            qs1 = "select * from Orders where O_way != '來店';";
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
                        reader["O_name"],reader["O_tel"],reader["O_address"],reader["O_stat"]});
                    }
                    reader.Close();
                }
            }
            
        }
        private void btn加入明細_Click(object sender, EventArgs e)
        {
            j = 0;
            try
            {
                if (int.TryParse(tb品項1數量.Text, out a) == true)
                {
                    if (dg門市暫存明細.RowCount != 0)
                    {
                        for (icount = 0; icount < dg門市暫存明細.RowCount; icount++)
                        {
                            if (cb品項1.SelectedItem.ToString() == dg門市暫存明細.Rows[icount].Cells[0].Value.ToString())
                            {
                                j++;
                                break;
                            }
                        }
                        if (j == 1)
                        {
                            dg門市暫存明細.Rows[icount].Cells[1].Value = Convert.ToInt32(dg門市暫存明細.Rows[icount].Cells[1].Value) + a;
                            dg門市暫存明細.Rows[icount].Cells[3].Value = Convert.ToInt32(dg門市暫存明細.Rows[icount].Cells[2].Value) *
                                Convert.ToInt32(dg門市暫存明細.Rows[icount].Cells[1].Value);
                        }
                        else
                        {
                            rows = dg門市暫存明細.Rows;
                            rows.Add(new object[] { cb品項1.SelectedItem.ToString(), tb品項1數量.Text, lb品項1單價.Text,
                            (Convert.ToInt32(lb品項1單價.Text) * Convert.ToInt32(tb品項1數量.Text)).ToString()});
                        }


                    }
                    else
                    {
                        rows = dg門市暫存明細.Rows;
                        rows.Add(new object[] { cb品項1.SelectedItem.ToString(), tb品項1數量.Text, lb品項1單價.Text,
                        (Convert.ToInt32(lb品項1單價.Text) * Convert.ToInt32(tb品項1數量.Text)).ToString()});
                    }
                    count = 0;
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
                else if (cb品項1.SelectedIndex < 0)
                {
                    MessageBox.Show("請先選擇商品");
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
            catch (Exception)
            {
                MessageBox.Show("資料有誤無法加入");
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

            if (indexrows > -1)
            {
                cb品項1.SelectedIndex = -1;
                tb品項1數量.Text = "0";
                lb庫存.Text = "0";
                lb品項1單價.Text = "0";
                lb品項1小計.Text = "0";
                if (dg門市暫存明細.RowCount > 1)
                {
                    dg門市暫存明細.Rows.RemoveAt(indexrows);
                }
            }
            else
            {
                MessageBox.Show("請先選取要刪除的品項");
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
            dg門市暫存明細.Rows.Clear();
        }
    }
}
