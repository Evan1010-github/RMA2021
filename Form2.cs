using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace RMA2021
{
    public partial class Form2 : Form
    {
        private readonly string ID;

        public Form2(string Model,string board,string SN,string ID)
        {
            InitializeComponent();
            label1.Text = Model;
            label2.Text = board;
            label3.Text = SN;
            this.ID = ID;
        }

        private void button1_Click(object sender, EventArgs e)
        {
           string connectionString = @"server=192.168.1.31;port=36288;userid=rma;password=GdUmm0J4EnJZneue;database=rma;charset=utf8";
            MySqlConnection conn = new MySqlConnection(connectionString);
            conn.Open();//開啟通道，建立連線，出現異常時,使用try catch語句
            using var cmd = new MySqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = "UPDATE rma.micfactoryrepair SET 需求日期='" + dateTimePicker1.Value.ToString("yyyy" + "MM" + "dd") +
                    "' WHERE 流水號='" + ID + "'";
            cmd.ExecuteNonQuery();
            conn.Close();
            MessageBox.Show("成功更新資料!");
            
        }
    }
}
