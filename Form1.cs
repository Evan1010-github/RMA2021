using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.Drawing.Printing;
using System.Data.SqlClient;
using RMA2021.Properties;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.Net;

namespace RMA2021
{
    public partial class Form1 : Form
    {
        // string connectionString = @"server=localhost;userid=root;password=1010;database=rma";
        private readonly string connectionString = @"server=192.168.1.31;port=36288;userid=rma;password=GdUmm0J4EnJZneue;database=rma;charset=utf8";
        int bookID = 0;
        int bookComID = 0;
        string txtBulidDate = "";
        string txtBulidPerson = "";
        string txtRepairPerson = "";
        string txtCalPerson = "";
        string txtFQCPerson = "";
        string PlayRole = "";
        //   bool warning = false;
        public Form1()
        {
            InitializeComponent();
            dataGridView1.Enabled = false;
            dataGridView2.Enabled = false;
            dataGridView3.Enabled = false;
            dataGridView4.Enabled = false;
            DGVcom.Enabled = false;
            ///業務才開(客訴區)
            txtComNumber.Enabled = false;
            txtComModel.Enabled = false;
            dateTimePickerCom.Enabled = false;
            txtComCustomer.Enabled = false;
            CBComWarranty.Enabled = false;
            txtComAppearance.Enabled = false;
            CBComAppearanceSort.Enabled = false;
            ///FQC才開(客訴區)
            txtComCause.Enabled = false;
            CBComCasueSort.Enabled = false;
            txtImprovement.Enabled = false;
            txtComPerson.Enabled = false;
            CBComDepartment.Enabled = false;
            dateTimePickerComFinish.Enabled = false;
            txtComImNow.Enabled = false;
            CBComCur.Enabled = false;
            ///
            groupBoxRMA.Enabled = false;
            groupBoxCom.Enabled = false;
            txtFixed.Enabled = false;
            textFactoryFixed.Enabled = false;
            textFactoryFixed.Enabled = false;
            txtOldFixed.Enabled = false;
            btnRepairFinish.Enabled = false;
            BtnFactoryFixed.Enabled = false;
            BtnFactoryFixed.Enabled = false;
            btnDoNotRepair.Enabled = false;
            btnTestFinish.Enabled = false;
            btnTestSentRepair.Enabled = false;
            btnFQCFinish.Enabled = false;
            btnFQCSentTest.Enabled = false;
            btnOutput.Enabled = false;
            btnFactoryOutput.Enabled = false;
            dataGridViewItemsQuery.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            RBtn1.Checked = true;
            labelCOUNT.Text = "";
            timer1.Enabled = false;
            // Create the ToolTip and associate with the Form container.
            ToolTip toolTip = new ToolTip();
            // Set up the delays for the ToolTip.
            toolTip.AutoPopDelay = 5000;    //工具提示保持可見的時間期限
            toolTip.InitialDelay = 100;     //滑鼠放上，自動開啟提示的時間
            toolTip.ReshowDelay = 50;       //滑鼠離開，自動關閉提示的時間
            toolTip.ShowAlways = true;     //總是顯示，即便空間非活動
            toolTip.UseAnimation = true;   //動畫效果
            toolTip.UseFading = true;      //淡入淡出效果
            toolTip.IsBalloon = true;      //氣球狀外觀
            // Set up the ToolTip text for the Button
            toolTip.SetToolTip(this.txtWarranty, "自動依建單時輸入的出貨日，以一年內計算");
            BTNDeadLine.Enabled = false;
            Clear();
            ClearCOM();
            GridFill();
            softUpdateCheck();
            radioButton1.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            radioButton2.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            radioButton3.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            radioButton4.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.ActiveControl = txtUserID;
        }
        private void softUpdateCheck()
        {
            linkLabel1.Text = "RMA更新";
            double RMAversion = 4.1;
            string DefultFormText = "RMA V4.1安裝版";//first Load Text
            // Connect to the MySQL database.
            string cs1 = @"server=192.168.1.31;port=36288;userid=rma;password=GdUmm0J4EnJZneue;database=rma;charset=utf8";
            using (MySqlConnection con = new MySqlConnection(cs1))
            {
                // Try to open the connection.
                try
                {
                    con.Open();

                    // Get the latest software version from the database.
                    string sql = "SELECT * FROM rma.SoftwareVersion;";
                    using (MySqlCommand cmd = new MySqlCommand(sql, con))
                    using (MySqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // Iterate through the results.
                        while (rdr.Read())
                        {
                            // Get the software name and version.
                            string softName = rdr["SoftName"].ToString();
                            double version = double.Parse(rdr["Version"].ToString());
                            if (softName == "RMA")
                            {
                                if (version > RMAversion)
                                {
                                    // The latest version is newer than the current version.
                                    this.Text = DefultFormText;
                                    linkLabel1.Text = "最新版為V" + version + "請按此更新";
                                }
                                else
                                {
                                    // The latest version is older than or equal to the current version.
                                    this.Text = DefultFormText;
                                    linkLabel1.Text = "當前版本V" + version + "已是最新版";
                                }
                            }
                        }
                    }
                }
                catch (SqlException ex)
                {
                    // Display an error message.
                    MessageBox.Show("MySQL SERVER連線異常!!!");
                    Console.WriteLine(ex.Message);
                }
            }
        }

        void Clear()
        {
            txtRepairOrCal.Text = txtFinishOrSemi.Text = txtClient.Text = txtSearch.Text = "";
            txtWarranty.Text = ""; txtModelName.Text = ""; txtBoardName.Text = ""; txtFinishSN.Text = "";
            txtSemiSN.Text = ""; txtBranch.Text = ""; txtSales.Text = ""; txtAccessories.Text = ""; txtOldFixed2.Text = "";
            txtVer.Text = ""; txtReturnCause.Text = ""; textBoxFaultCause.Text = ""; txtFixed.Text = ""; txtOldFixed.Text = ""; textFactoryFixed.Text = "";
            txtFinishMark.Text = ""; textVolt.Text = ""; dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = " ";// dateTimePicker1.Text = NULL;
            bookID = 0;
            btnSave.Text = "輸入";
            txtStatus.Text = "";
            txtFinishMark.Enabled = false;
            TBBuildDay.Text = ""; TBFixedDay.Text = ""; TBTestedDay.Text = ""; TBFQCFinish.Text = ""; TBCloseDay.Text = "";
            txtUserNameShow.Enabled = false;
            txtUserNameShow.Items.Clear();
        }
        private void dateTimePicker1_ValueChanged(object sender, System.EventArgs e)
        {
            DateTime dt1 = dateTimePicker1.Value;
            DateTime dt2 = DateTime.Now.AddDays(-365);
            int n = dt1.CompareTo(dt2);
            this.dateTimePicker1.Format = DateTimePickerFormat.Long;
            this.dateTimePicker1.CustomFormat = null;
            if (n < 0)
            {//過保
                txtWarranty.Text = "否";
            }
            else
            {//未過保
                txtWarranty.Text = "是";
            }
        }
        void CalDays()
        {
            DateTime date_start = Convert.ToDateTime(TBBuildDay.Text);
            DateTime date_today = DateTime.Now.ToLocalTime();
            TimeSpan Diff_datetoday = date_today.Subtract(date_start);
            textBoxtoday.Text = Diff_datetoday.Days.ToString();
            if (TBFixedDay.Text != "")
            {
                DateTime date_1 = Convert.ToDateTime(TBBuildDay.Text);
                DateTime date_2 = Convert.ToDateTime(TBFixedDay.Text);
                TimeSpan Diff_dates = date_2.Subtract(date_1);
                textBox21.Text = Diff_dates.Days.ToString();
            }
            else
            {
                textBox21.Text = "";
            }
            if (TBTestedDay.Text != "" & TBFixedDay.Text != "")
            {
                DateTime date_2 = Convert.ToDateTime(TBFixedDay.Text);
                DateTime date_3 = Convert.ToDateTime(TBTestedDay.Text);
                TimeSpan Diff_dates = date_3.Subtract(date_2);
                textBox32.Text = Diff_dates.Days.ToString();
            }
            else
            {
                textBox32.Text = "";
            }
            if (TBFQCFinish.Text != "" & TBTestedDay.Text != "")
            {
                DateTime date_4 = Convert.ToDateTime(TBFQCFinish.Text);
                DateTime date_3 = Convert.ToDateTime(TBTestedDay.Text);
                TimeSpan Diff_dates = date_4.Subtract(date_3);
                textBox43.Text = Diff_dates.Days.ToString();
            }
            else
            {
                textBox43.Text = "";
            }
            if (TBFQCFinish.Text != "" & TBCloseDay.Text != "")
            {
                DateTime date_4 = Convert.ToDateTime(TBFQCFinish.Text);
                DateTime date_5 = Convert.ToDateTime(TBCloseDay.Text);
                TimeSpan Diff_dates = date_5.Subtract(date_4);
                textBox54.Text = Diff_dates.Days.ToString();
            }
            else
            {
                textBox54.Text = "";
            }
            if (TBCloseDay.Text != "")
            {
                DateTime date_1 = Convert.ToDateTime(TBBuildDay.Text);
                DateTime date_5 = Convert.ToDateTime(TBCloseDay.Text);
                TimeSpan Diff_dates = date_5.Subtract(date_1);
                textBox51.Text = Diff_dates.Days.ToString();
            }
            else
            {
                textBox51.Text = "";
            }
        }

        void GridFill()
        {
            try
            {
                using MySqlConnection mysqlCon = new MySqlConnection(connectionString);
                mysqlCon.Open();
                MySqlDataAdapter sqlDa = new MySqlDataAdapter("RMAViewByRepair", mysqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                System.Data.DataTable dtblBook = new System.Data.DataTable();
                sqlDa.Fill(dtblBook);
                dataGridView1.DataSource = dtblBook;
                dataGridView1.DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView1.DefaultCellStyle.BackColor = Color.Beige;
                dataGridView1.Columns[0].Visible = false;//狀態txtStatus
                dataGridView1.Columns[1].Visible = false;//流水號bookID
                dataGridView1.Columns[2].Visible = false;//維修/校驗txtRepairOrCal
                dataGridView1.Columns[3].Visible = false;//成品/半成品txtFinishOrSemi
                                                         //dataGridView1.Columns[4].Visible = false;//機種名txtModelName
                                                         // dataGridView1.Columns[5].Visible = false;//板名txtBoardName
                                                         // dataGridView1.Columns[6].Visible = false;//成品序號txtFinishSN
                                                         // dataGridView1.Columns[7].Visible = false;//半成品序號txtSemiSN
                dataGridView1.Columns[8].Visible = false;//送件據點txtBranch
                dataGridView1.Columns[9].Visible = false;//所屬業務txtSales
                dataGridView1.Columns[10].Visible = false;//保固內txtWarranty
                                                          // dataGridView1.Columns[11].Visible = false;//客戶名txtClient
                dataGridView1.Columns[12].Visible = false;//配件txtAccessories
                dataGridView1.Columns[13].Visible = false;//版本txtVer
                dataGridView1.Columns[14].Visible = false;//故障描述txtReturnCause
                                                          // dataGridView1.Columns[15].Visible = false;//建單日txtBulidDate
                                                          // dataGridView1.Columns[16].Visible = false;//建單人txtBulidPerson
                dataGridView1.Columns[17].Visible = false;//維修內容txtFixed
                dataGridView1.Columns[18].Visible = false;//維修完成日txtRepairFinD
                dataGridView1.Columns[19].Visible = false;//維修人員txtRepairPerson
                dataGridView1.Columns[20].Visible = false;//測試完成日txtCalFinD
                dataGridView1.Columns[21].Visible = false;//測試人員txtCalPerson
                dataGridView1.Columns[22].Visible = false;//FQC完成日txtFQCFinD
                dataGridView1.Columns[23].Visible = false;//FQC人員txtFQCPerson
                dataGridView1.Columns[24].Visible = false;//結案完成日txtCaseCloseD
                dataGridView1.Columns[25].Visible = false;//結案人員txtCaseClosePerson
                dataGridView1.Columns[26].Visible = false;//測試故障描述

                dataGridView1.Columns[28].Visible = true;//出貨日dateTimepick                  //2020/2/8  added
                dataGridView1.Columns[29].Visible = false;//使用電壓texeVolt                    //2020/2/8  added

                MySqlDataAdapter sqlDa2 = new MySqlDataAdapter("RMAViewByCal", mysqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                System.Data.DataTable dtblBook2 = new System.Data.DataTable();
                sqlDa2.Fill(dtblBook2);
                dataGridView2.DataSource = dtblBook2;
                dataGridView2.DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView2.DefaultCellStyle.BackColor = Color.Beige;
                dataGridView2.Columns[0].Visible = false;//狀態txtStatus
                dataGridView2.Columns[1].Visible = false;//流水號bookID
                dataGridView2.Columns[2].Visible = false;//維修/校驗txtRepairOrCal
                dataGridView2.Columns[3].Visible = false;//成品/半成品txtFinishOrSemi
                                                         //dataGridView2.Columns[4].Visible = false;//機種名txtModelName
                                                         // dataGridView2.Columns[5].Visible = false;//板名txtBoardName
                                                         // dataGridView2.Columns[6].Visible = false;//成品序號txtFinishSN
                                                         // dataGridView2.Columns[7].Visible = false;//半成品序號txtSemiSN
                dataGridView2.Columns[8].Visible = false;//送件據點txtBranch
                dataGridView2.Columns[9].Visible = false;//所屬業務txtSales
                dataGridView2.Columns[10].Visible = false;//保固內txtWarranty
                                                          // dataGridView2.Columns[11].Visible = false;//客戶名txtClient
                dataGridView2.Columns[12].Visible = false;//配件txtAccessories
                dataGridView2.Columns[13].Visible = false;//版本txtVer
                dataGridView2.Columns[14].Visible = false;//故障描述txtReturnCause
                dataGridView2.Columns[15].Visible = false;//建單日txtBulidDate
                dataGridView2.Columns[16].Visible = false;//建單人txtBulidPerson
                dataGridView2.Columns[17].Visible = false;//維修內容txtFixed
                                                          //dataGridView2.Columns[18].Visible = false;//維修完成日txtRepairFinD
                                                          // dataGridView2.Columns[19].Visible = false;//維修人員txtRepairPerson
                dataGridView2.Columns[20].Visible = false;//測試完成日txtCalFinD
                dataGridView2.Columns[21].Visible = false;//測試人員txtCalPerson
                dataGridView2.Columns[22].Visible = false;//FQC完成日txtFQCFinD
                dataGridView2.Columns[23].Visible = false;//FQC人員txtFQCPerson
                dataGridView2.Columns[24].Visible = false;//結案完成日txtCaseCloseD
                dataGridView2.Columns[25].Visible = false;//結案人員txtCaseClosePerson

                dataGridView2.Columns[28].Visible = false;//出貨日dateTimepick                  //2020/2/8  added
                dataGridView2.Columns[29].Visible = false;//使用電壓texeVolt                    //2020/2/8  added
                MySqlDataAdapter sqlDa3 = new MySqlDataAdapter("RMAViewByFQC", mysqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                System.Data.DataTable dtblBook3 = new System.Data.DataTable();
                sqlDa3.Fill(dtblBook3);
                dataGridView3.DataSource = dtblBook3;
                dataGridView3.DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView3.DefaultCellStyle.BackColor = Color.Beige;
                dataGridView3.Columns[0].Visible = false;//狀態txtStatus
                dataGridView3.Columns[1].Visible = false;//流水號bookID
                dataGridView3.Columns[2].Visible = false;//維修/校驗txtRepairOrCal
                dataGridView3.Columns[3].Visible = false;//成品/半成品txtFinishOrSemi
                                                         //dataGridView3.Columns[4].Visible = false;//機種名txtModelName
                                                         // dataGridView3.Columns[5].Visible = false;//板名txtBoardName
                                                         // dataGridView3.Columns[6].Visible = false;//成品序號txtFinishSN
                                                         // dataGridView3.Columns[7].Visible = false;//半成品序號txtSemiSN
                dataGridView3.Columns[8].Visible = false;//送件據點txtBranch
                dataGridView3.Columns[9].Visible = false;//所屬業務txtSales
                dataGridView3.Columns[10].Visible = false;//保固內txtWarranty
                                                          // dataGridView3.Columns[11].Visible = false;//客戶名txtClient
                dataGridView3.Columns[12].Visible = false;//配件txtAccessories
                dataGridView3.Columns[13].Visible = false;//版本txtVer
                dataGridView3.Columns[14].Visible = false;//故障描述txtReturnCause
                dataGridView3.Columns[15].Visible = false;//建單日txtBulidDate
                dataGridView3.Columns[16].Visible = false;//建單人txtBulidPerson
                dataGridView3.Columns[17].Visible = false;//維修內容txtFixed
                dataGridView3.Columns[18].Visible = false;//維修完成日txtRepairFinD
                dataGridView3.Columns[19].Visible = false;//維修人員txtRepairPerson
                                                          // dataGridView3.Columns[20].Visible = false;//測試完成日txtCalFinD
                                                          // dataGridView3.Columns[21].Visible = false;//測試人員txtCalPerson
                dataGridView3.Columns[22].Visible = false;//FQC完成日txtFQCFinD
                dataGridView3.Columns[23].Visible = false;//FQC人員txtFQCPerson
                dataGridView3.Columns[24].Visible = false;//結案完成日txtCaseCloseD
                dataGridView3.Columns[25].Visible = false;//結案人員txtCaseClosePerson

                dataGridView3.Columns[28].Visible = false;//出貨日dateTimepick                  //2020/2/8  added
                dataGridView3.Columns[29].Visible = false;//使用電壓texeVolt                    //2020/2/8  added
                MySqlDataAdapter sqlDa4 = new MySqlDataAdapter("RMAViewByFinish", mysqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                System.Data.DataTable dtblBook4 = new System.Data.DataTable();
                sqlDa4.Fill(dtblBook4);
                dataGridView4.DataSource = dtblBook4;
                dataGridView4.DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView4.DefaultCellStyle.BackColor = Color.Beige;
                dataGridView4.Columns[0].Visible = false;//狀態txtStatus
                dataGridView4.Columns[1].Visible = false;//流水號bookID
                dataGridView4.Columns[2].Visible = false;//維修/校驗txtRepairOrCal
                dataGridView4.Columns[3].Visible = false;//成品/半成品txtFinishOrSemi
                                                         //dataGridView4.Columns[4].Visible = false;//機種名txtModelName
                                                         // dataGridView4.Columns[5].Visible = false;//板名txtBoardName
                                                         // dataGridView4.Columns[6].Visible = false;//成品序號txtFinishSN
                                                         // dataGridView4.Columns[7].Visible = false;//半成品序號txtSemiSN
                dataGridView4.Columns[8].Visible = false;//送件據點txtBranch
                dataGridView4.Columns[9].Visible = false;//所屬業務txtSales
                dataGridView4.Columns[10].Visible = false;//保固內txtWarranty
                                                          // dataGridView4.Columns[11].Visible = false;//客戶名txtClient
                dataGridView4.Columns[12].Visible = false;//配件txtAccessories
                dataGridView4.Columns[13].Visible = false;//版本txtVer
                dataGridView4.Columns[14].Visible = false;//故障描述txtReturnCause
                dataGridView4.Columns[15].Visible = false;//建單日txtBulidDate
                dataGridView4.Columns[16].Visible = false;//建單人txtBulidPerson
                dataGridView4.Columns[17].Visible = false;//維修內容txtFixed
                dataGridView4.Columns[18].Visible = false;//維修完成日txtRepairFinD
                dataGridView4.Columns[19].Visible = false;//維修人員txtRepairPerson
                dataGridView4.Columns[20].Visible = false;//測試完成日txtCalFinD
                dataGridView4.Columns[21].Visible = false;//測試人員txtCalPerson
                                                          // dataGridView4.Columns[22].Visible = false;//FQC完成日txtFQCFinD
                                                          // dataGridView4.Columns[23].Visible = false;//FQC人員txtFQCPerson
                dataGridView4.Columns[24].Visible = false;//結案完成日txtCaseCloseD
                dataGridView4.Columns[25].Visible = false;//結案人員txtCaseClosePerson

                dataGridView4.Columns[28].Visible = false;//出貨日dateTimepick                  //2020/2/8  added
                dataGridView4.Columns[29].Visible = false;//使用電壓texeVolt                    //2020/2/8  added
                MySqlDataAdapter sqlDaMIC = new MySqlDataAdapter("MICFactoryViewByRepair", mysqlCon);//待維修
                sqlDaMIC.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable dtblBookMIC = new DataTable();
                sqlDaMIC.Fill(dtblBookMIC);
                DGVFactoryRepair.DataSource = dtblBookMIC;
                DGVFactoryRepair.DefaultCellStyle.ForeColor = Color.Blue;
                DGVFactoryRepair.DefaultCellStyle.BackColor = Color.Beige;
                DGVFactoryRepair.Columns[0].Visible = false;
                DGVFactoryRepair.Columns[12].Visible = false;
                DGVFactoryRepair.Columns[13].Visible = false;
                DGVFactoryRepair.Columns[14].Visible = false;
                // dataGridViewMICRepair.Columns[15].Visible = false;
                MySqlDataAdapter sqlDaMIC2 = new MySqlDataAdapter("MICFactoryViewByRepaired", mysqlCon);//維修完成
                sqlDaMIC.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable dtblBookMIC2 = new DataTable();
                sqlDaMIC2.Fill(dtblBookMIC2);
                DGVFactoryFixed.DataSource = dtblBookMIC2;
                DGVFactoryFixed.DefaultCellStyle.ForeColor = Color.Blue;
                DGVFactoryFixed.DefaultCellStyle.BackColor = Color.Beige;
                DGVFactoryFixed.Columns[0].Visible = false;
                DGVFactoryFixed.Columns[8].Visible = false;
                DGVFactoryFixed.Columns[9].Visible = false;
                DGVFactoryFixed.Columns[10].Visible = false;
                DGVFactoryFixed.Columns[11].Visible = false;
                //DGVFactoryFixed.Columns[15].Visible = false;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("SQL SERVER連線異常!!!");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //mysqlCon.Close();
            }
        }
        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 || e.KeyCode == Keys.Enter) //<---判斷是否按下Enter
                                                             //  if (e.KeyCode == Keys.Enter)//<---判斷是否按下Enter
            {
                btnLogin_Click(this, null);
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            txtReturnCause.Text = txtReturnCause.Text.Replace("'", "\\'");//2022/11/10 can't 輸入報錯問題
            if (txtStatus.Text != "待結案")
            {
                if (txtRepairOrCal.Text == "維修")//Sales輸入時Status
                {
                    txtStatus.Text = "待維修";
                }
                if (txtRepairOrCal.Text == "校驗")
                {
                    txtStatus.Text = "待校驗";
                }
            }
            txtBulidDate = DateTime.Now.ToLocalTime().ToString();//建單日時
            ////SQL  開始write //////////////////////////////////////////////////////////////////////////////////////////////    
            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();//開啟通道，建立連線，出現異常時,使用try catch語句
                using var cmd = new MySqlCommand();
                cmd.Connection = conn;
                if (bookID == 0)//新增資料
                {
                    txtBulidDate = DateTime.Now.ToLocalTime().ToString();
                    if (txtWarranty.Text == "") //2022/02/18   保固為空
                    {
                        cmd.CommandText = $"INSERT INTO rma.rmarawdata(狀態,維修校驗,成品半成品,機種名," +
                       $"板名,成品序號,半成品序號,送件據點,所屬業務,保固內,客戶名,配件,版本,故障描述,建單日,出貨日,使用電壓,建單人)" +
                               // $",測試完成日,測試人員,FQC完成日,FQC人員,結案完成日,結案人員)" +
                               "VALUES('" + txtStatus.Text + "'" +
                               ",'" + txtRepairOrCal.Text + "'" +
                               ",'" + txtFinishOrSemi.Text + "'" +
                               ",'" + txtModelName.Text + "'" +
                               ",'" + txtBoardName.Text + "'" +
                               ",'" + txtFinishSN.Text + "'" +
                               ",'" + txtSemiSN.Text + "'" +
                               ",'" + txtBranch.Text + "'" +
                               ",'" + txtSales.Text + "'" +
                               ",'" + "N/A" + "'" +
                               ",'" + txtClient.Text + "'" +
                               ",'" + txtAccessories.Text + "'" +
                               ",'" + txtVer.Text + "'" +
                               ",'" + txtReturnCause.Text + "'" +
                               ",'" + txtBulidDate + "'" +
                                ",'" + "" + "'" +    //2022/2/8  added
                               ",'" + textVolt.Text + "'" +            //
                               ",'" + txtBulidPerson + "')";
                    }
                    else
                    {
                        cmd.CommandText = $"INSERT INTO rma.rmarawdata(狀態,維修校驗,成品半成品,機種名," +
                            $"板名,成品序號,半成品序號,送件據點,所屬業務,保固內,客戶名,配件,版本,故障描述,建單日,出貨日,使用電壓,建單人)" +
                                    // $",測試完成日,測試人員,FQC完成日,FQC人員,結案完成日,結案人員)" +
                                    "VALUES('" + txtStatus.Text + "'" +
                                    ",'" + txtRepairOrCal.Text + "'" +
                                    ",'" + txtFinishOrSemi.Text + "'" +
                                    ",'" + txtModelName.Text + "'" +
                                    ",'" + txtBoardName.Text + "'" +
                                    ",'" + txtFinishSN.Text + "'" +
                                    ",'" + txtSemiSN.Text + "'" +
                                    ",'" + txtBranch.Text + "'" +
                                    ",'" + txtSales.Text + "'" +
                                    ",'" + txtWarranty.Text + "'" +
                                    ",'" + txtClient.Text + "'" +
                                    ",'" + txtAccessories.Text + "'" +
                                    ",'" + txtVer.Text + "'" +
                                    ",'" + txtReturnCause.Text + "'" +
                                    ",'" + txtBulidDate + "'" +
                                     ",'" + dateTimePicker1.Text + "'" +    //2022/2/8  added
                                    ",'" + textVolt.Text + "'" +            //
                                    ",'" + txtBulidPerson + "')";
                    }
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("成功輸入資料!");
                }
                else//更新資料，只有建立者能更新or刪除
                {
                    switch (PlayRole)
                    {
                        case "sales":
                            if (txtStatus.Text == "待結案")
                            {
                                cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + "已結案" +
                                "',結案完成日='" + DateTime.Now.ToLocalTime().ToString() +
                                "',結案人員 ='" + txtBulidPerson +
                                "',結案描述 ='" + txtFinishMark.Text +
                                 "' WHERE 流水號='" + bookID + "'";
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                            else//更新
                            {
                                if (txtWarranty.Text == "") //2022/02/18   保固為空
                                {
                                    cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + txtStatus.Text +
                                 "',維修校驗='" + txtRepairOrCal.Text +
                                  "',成品半成品='" + txtFinishOrSemi.Text +
                                  "',機種名='" + txtModelName.Text +
                                   "',板名='" + txtBoardName.Text +
                                  "',成品序號='" + txtFinishSN.Text +
                                  "',半成品序號='" + txtSemiSN.Text +
                                  "',送件據點='" + txtBranch.Text +
                                  "',所屬業務='" + txtSales.Text +
                                  "',保固內='" + "N/A" +
                                  "',客戶名='" + txtClient.Text +
                                 "',配件='" + txtAccessories.Text +
                                 "',版本='" + txtVer.Text +
                             "',故障描述='" + "@" + "\"" + txtReturnCause.Text + "\"" + //  可逃逸? 
                                                                                    //         "',故障描述='" + txtReturnCause.Text +
                                 "',出貨日='" + "" +    //2022/2/8  added
                                  "',使用電壓='" + textVolt.Text +         //2022/2/8  added
                                 "',機種名='" + txtModelName.Text +
                                 "' WHERE 流水號='" + bookID + "'";
                                }
                                else
                                {
                                    cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + txtStatus.Text +
                                 "',維修校驗='" + txtRepairOrCal.Text +
                                  "',成品半成品='" + txtFinishOrSemi.Text +
                                  "',機種名='" + txtModelName.Text +
                                   "',板名='" + txtBoardName.Text +
                                  "',成品序號='" + txtFinishSN.Text +
                                  "',半成品序號='" + txtSemiSN.Text +
                                  "',送件據點='" + txtBranch.Text +
                                  "',所屬業務='" + txtSales.Text +
                                  "',保固內='" + txtWarranty.Text +
                                  "',客戶名='" + txtClient.Text +
                                 "',配件='" + txtAccessories.Text +
                                 "',版本='" + txtVer.Text +
                                 "',故障描述='" + txtReturnCause.Text +
                                 "',出貨日='" + dateTimePicker1.Text +    //2022/2/8  added
                                  "',使用電壓='" + textVolt.Text +         //2022/2/8  added
                                 "',機種名='" + txtModelName.Text +
                                 "' WHERE 流水號='" + bookID + "'";
                                }
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                            break;
                        case "eng":
                            cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + "待校驗" +
                            "',維修內容='" + txtOldFixed.Text + "#" + txtFixed.Text +
                            "',維修完成日='" + DateTime.Now.ToLocalTime().ToString() +
                            "',維修人員 ='" + txtRepairPerson +
                             "' WHERE 流水號='" + bookID + "'";
                            cmd.ExecuteNonQuery();
                            conn.Close();
                            break;

                        case "test":
                            cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + "待FQC" +
                            "',測試完成日='" + DateTime.Now.ToLocalTime().ToString() +
                            "',測試人員 ='" + txtCalPerson +
                             "' WHERE 流水號='" + bookID + "'";
                            cmd.ExecuteNonQuery();
                            conn.Close();

                            break;
                        case "FQC":
                            cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + "待結案" +
                            "',FQC完成日='" + DateTime.Now.ToLocalTime().ToString() +
                            "',FQC人員 ='" + txtFQCPerson +
                             "' WHERE 流水號='" + bookID + "'";
                            cmd.ExecuteNonQuery();
                            conn.Close();

                            break;
                        default:
                            conn.Close();

                            break;
                    }
                    MessageBox.Show("成功更新資料!");
                }
                Clear();
                GridFill();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("save SERVER連線異常!!!");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void dgvBook_DoubleClick(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            btnDelete.Enabled = false;
            if (dataGridView1.CurrentRow.Index != -1)
            {
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                txtStatus.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                txtRepairOrCal.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                txtFinishOrSemi.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                txtModelName.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                txtBoardName.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                txtFinishSN.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                txtSemiSN.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                txtBranch.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                txtSales.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                txtWarranty.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                txtClient.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                txtAccessories.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                txtVer.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                txtReturnCause.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                txtFinishMark.Text = "";
                // txtBulidDate = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                txtOldFixed.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                txtOldFixed2.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                // dataGridView1.Rows[dataGridView1.CurrentRow.Index].DefaultCellStyle.BackColor = Color.Yellow;
                TBBuildDay.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();//建單日
                TBFixedDay.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();//維修完成日txtRepairFinD
                TBTestedDay.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();//測試完成日txtCalFinD
                TBFQCFinish.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();//FQC完成日txtFQCFinD
                TBCloseDay.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();//結案完成日txtCaseCloseD
                textBoxFaultCause.Text = dataGridView1.CurrentRow.Cells[27].Value.ToString();//測試不良描述

                //2020/2/8  added
                if (dataGridView1.CurrentRow.Cells[28].Value.ToString() == "" || dataGridView1.CurrentRow.Cells[28].Value.ToString() == " ")
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = " ";// dateTimePicker1.Text = NULL;
                }
                else
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Long;
                    dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[28].Value.ToString();//出貨日dateTimepick 
                }
                textVolt.Text = dataGridView1.CurrentRow.Cells[29].Value.ToString();//使用電壓texeVolt                    //2020/2/8  added
                CalDays();
            }
            bookID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            if (txtUserNameShow.Text == dataGridView1.CurrentRow.Cells[16].Value.ToString())//非本人建立不得修改or delete
            {
                btnSave.Text = "更新";
                btnSave.Enabled = true;
                btnDelete.Enabled = true;
            }
            if (PlayRole == "eng")//維修人員
            {
                btnSave.Text = "更新";
                btnSave.Enabled = true;
            }
        }
        private void dgvBook_DoubleClick2(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            btnDelete.Enabled = false;
            if (dataGridView2.CurrentRow.Index != -1)
            {
                dataGridView2.Rows[dataGridView2.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                dataGridView2.Rows[dataGridView2.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                txtStatus.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
                txtRepairOrCal.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                txtFinishOrSemi.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
                txtModelName.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();
                txtBoardName.Text = dataGridView2.CurrentRow.Cells[5].Value.ToString();
                txtFinishSN.Text = dataGridView2.CurrentRow.Cells[6].Value.ToString();
                txtSemiSN.Text = dataGridView2.CurrentRow.Cells[7].Value.ToString();
                txtBranch.Text = dataGridView2.CurrentRow.Cells[8].Value.ToString();
                txtSales.Text = dataGridView2.CurrentRow.Cells[9].Value.ToString();
                txtWarranty.Text = dataGridView2.CurrentRow.Cells[10].Value.ToString();
                txtClient.Text = dataGridView2.CurrentRow.Cells[11].Value.ToString();
                txtAccessories.Text = dataGridView2.CurrentRow.Cells[12].Value.ToString();
                txtVer.Text = dataGridView2.CurrentRow.Cells[13].Value.ToString();
                txtReturnCause.Text = dataGridView2.CurrentRow.Cells[14].Value.ToString();
                txtFinishMark.Text = "";
                // txtBulidDate = dataGridView2.CurrentRow.Cells[15].Value.ToString();
                // dataGridView1.Rows[dataGridView2.CurrentRow.Index].DefaultCellStyle.BackColor = Color.Yellow;
                TBBuildDay.Text = dataGridView2.CurrentRow.Cells[15].Value.ToString();//建單日
                TBFixedDay.Text = dataGridView2.CurrentRow.Cells[18].Value.ToString();//維修完成日txtRepairFinD
                TBTestedDay.Text = dataGridView2.CurrentRow.Cells[20].Value.ToString();//測試完成日txtCalFinD
                TBFQCFinish.Text = dataGridView2.CurrentRow.Cells[22].Value.ToString();//FQC完成日txtFQCFinD
                TBCloseDay.Text = dataGridView2.CurrentRow.Cells[24].Value.ToString();//結案完成日txtCaseCloseD
                txtOldFixed.Text = dataGridView2.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                txtOldFixed2.Text = dataGridView2.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                //2020/2/8  added
                if (dataGridView2.CurrentRow.Cells[28].Value.ToString() == "" || dataGridView2.CurrentRow.Cells[28].Value.ToString() == " ")
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = " ";// dateTimePicker1.Text = NULL;
                }
                else
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Long;
                    dateTimePicker1.Text = dataGridView2.CurrentRow.Cells[28].Value.ToString();//出貨日dateTimepick 
                }
                textVolt.Text = dataGridView2.CurrentRow.Cells[29].Value.ToString();//使用電壓texeVolt                    //2020/2/8  added
                CalDays();
            }
            bookID = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value.ToString());
            if (txtUserNameShow.Text == dataGridView2.CurrentRow.Cells[16].Value.ToString())//非本人建立不得修改or delete
            {
                btnSave.Text = "更新";
                btnSave.Enabled = true;
                btnDelete.Enabled = true;
            }
            if (PlayRole == "test")//測試人員
            {
                btnSave.Text = "測試完成";
                btnSave.Enabled = true;
            }
        }
        private void dgvBook_DoubleClick3(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            btnDelete.Enabled = false;
            if (dataGridView3.CurrentRow.Index != -1)
            {
                dataGridView3.Rows[dataGridView3.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                dataGridView3.Rows[dataGridView3.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                txtStatus.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
                txtRepairOrCal.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
                txtFinishOrSemi.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
                txtModelName.Text = dataGridView3.CurrentRow.Cells[4].Value.ToString();
                txtBoardName.Text = dataGridView3.CurrentRow.Cells[5].Value.ToString();
                txtFinishSN.Text = dataGridView3.CurrentRow.Cells[6].Value.ToString();
                txtSemiSN.Text = dataGridView3.CurrentRow.Cells[7].Value.ToString();
                txtBranch.Text = dataGridView3.CurrentRow.Cells[8].Value.ToString();
                txtSales.Text = dataGridView3.CurrentRow.Cells[9].Value.ToString();
                txtWarranty.Text = dataGridView3.CurrentRow.Cells[10].Value.ToString();
                txtClient.Text = dataGridView3.CurrentRow.Cells[11].Value.ToString();
                txtAccessories.Text = dataGridView3.CurrentRow.Cells[12].Value.ToString();
                txtVer.Text = dataGridView3.CurrentRow.Cells[13].Value.ToString();
                txtReturnCause.Text = dataGridView3.CurrentRow.Cells[14].Value.ToString();
                txtFinishMark.Text = "";
                //txtBulidDate = dataGridView3.CurrentRow.Cells[15].Value.ToString();
                // dataGridView1.Rows[dataGridView3.CurrentRow.Index].DefaultCellStyle.BackColor = Color.Yellow;
                TBBuildDay.Text = dataGridView3.CurrentRow.Cells[15].Value.ToString();//建單日
                TBFixedDay.Text = dataGridView3.CurrentRow.Cells[18].Value.ToString();//維修完成日txtRepairFinD
                TBTestedDay.Text = dataGridView3.CurrentRow.Cells[20].Value.ToString();//測試完成日txtCalFinD
                TBFQCFinish.Text = dataGridView3.CurrentRow.Cells[22].Value.ToString();//FQC完成日txtFQCFinD
                TBCloseDay.Text = dataGridView3.CurrentRow.Cells[24].Value.ToString();//結案完成日txtCaseCloseD
                txtOldFixed.Text = dataGridView3.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                txtOldFixed2.Text = dataGridView3.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                //2020/2/8  added
                if (dataGridView3.CurrentRow.Cells[28].Value.ToString() == "" || dataGridView3.CurrentRow.Cells[28].Value.ToString() == " ")
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = " ";// dateTimePicker1.Text = NULL;
                }
                else
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Long;
                    dateTimePicker1.Text = dataGridView3.CurrentRow.Cells[28].Value.ToString();//出貨日dateTimepick 
                }
                textVolt.Text = dataGridView3.CurrentRow.Cells[29].Value.ToString();//使用電壓texeVolt                    //2020/2/8  added
                CalDays();
            }
            bookID = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value.ToString());
            if (PlayRole == "FQC")//fqc人員
            {
                btnSave.Text = "FQC完成";
                btnSave.Enabled = true;
            }
        }
        private void dgvBook_DoubleClick4(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            btnDelete.Enabled = false;
            txtFinishMark.Enabled = false;
            if (dataGridView4.CurrentRow.Index != -1)
            {
                dataGridView4.Rows[dataGridView4.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                dataGridView4.Rows[dataGridView4.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                txtStatus.Text = dataGridView4.CurrentRow.Cells[1].Value.ToString();
                txtRepairOrCal.Text = dataGridView4.CurrentRow.Cells[2].Value.ToString();
                txtFinishOrSemi.Text = dataGridView4.CurrentRow.Cells[3].Value.ToString();
                txtModelName.Text = dataGridView4.CurrentRow.Cells[4].Value.ToString();
                txtBoardName.Text = dataGridView4.CurrentRow.Cells[5].Value.ToString();
                txtFinishSN.Text = dataGridView4.CurrentRow.Cells[6].Value.ToString();
                txtSemiSN.Text = dataGridView4.CurrentRow.Cells[7].Value.ToString();
                txtBranch.Text = dataGridView4.CurrentRow.Cells[8].Value.ToString();
                txtSales.Text = dataGridView4.CurrentRow.Cells[9].Value.ToString();
                txtWarranty.Text = dataGridView4.CurrentRow.Cells[10].Value.ToString();
                txtClient.Text = dataGridView4.CurrentRow.Cells[11].Value.ToString();
                txtAccessories.Text = dataGridView4.CurrentRow.Cells[12].Value.ToString();
                txtVer.Text = dataGridView4.CurrentRow.Cells[13].Value.ToString();
                txtReturnCause.Text = dataGridView4.CurrentRow.Cells[14].Value.ToString();
                txtFinishMark.Text = "";
                // txtBulidDate = dataGridView4.CurrentRow.Cells[15].Value.ToString();
                // dataGridView1.Rows[dataGridView3.CurrentRow.Index].DefaultCellStyle.BackColor = Color.Yellow;
                TBBuildDay.Text = dataGridView4.CurrentRow.Cells[15].Value.ToString();//建單日
                TBFixedDay.Text = dataGridView4.CurrentRow.Cells[18].Value.ToString();//維修完成日txtRepairFinD
                TBTestedDay.Text = dataGridView4.CurrentRow.Cells[20].Value.ToString();//測試完成日txtCalFinD
                TBFQCFinish.Text = dataGridView4.CurrentRow.Cells[22].Value.ToString();//FQC完成日txtFQCFinD
                TBCloseDay.Text = dataGridView4.CurrentRow.Cells[24].Value.ToString();//結案完成日txtCaseCloseD
                txtOldFixed.Text = dataGridView4.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                txtOldFixed2.Text = dataGridView4.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                //2020/2/8  added
                if (dataGridView4.CurrentRow.Cells[28].Value.ToString() == "" || dataGridView4.CurrentRow.Cells[28].Value.ToString() == " ")
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = " ";// dateTimePicker1.Text = NULL;
                }
                else
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Long;
                    dateTimePicker1.Text = dataGridView4.CurrentRow.Cells[28].Value.ToString();//出貨日dateTimepick 
                }
                textVolt.Text = dataGridView4.CurrentRow.Cells[29].Value.ToString();//使用電壓texeVolt                    //2020/2/8  added
                CalDays();
            }
            bookID = Convert.ToInt32(dataGridView4.CurrentRow.Cells[0].Value.ToString());
            if (txtUserNameShow.Text == dataGridView4.CurrentRow.Cells[16].Value.ToString())//非本人建立不得修改or delete
            {
                btnSave.Text = "結案";
                btnSave.Enabled = true;
                btnDelete.Enabled = false;
                txtFinishMark.Enabled = true;
            }
        }
        private void dgvBook_DoubleClick5(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            btnDelete.Enabled = false;
            if (dataGridView5.CurrentRow.Index != -1)
            {
                dataGridView5.Rows[dataGridView5.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                dataGridView5.Rows[dataGridView5.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                txtStatus.Text = dataGridView5.CurrentRow.Cells[1].Value.ToString();
                txtRepairOrCal.Text = dataGridView5.CurrentRow.Cells[2].Value.ToString();
                txtFinishOrSemi.Text = dataGridView5.CurrentRow.Cells[3].Value.ToString();
                txtModelName.Text = dataGridView5.CurrentRow.Cells[4].Value.ToString();
                txtBoardName.Text = dataGridView5.CurrentRow.Cells[5].Value.ToString();
                txtFinishSN.Text = dataGridView5.CurrentRow.Cells[6].Value.ToString();
                txtSemiSN.Text = dataGridView5.CurrentRow.Cells[7].Value.ToString();
                txtBranch.Text = dataGridView5.CurrentRow.Cells[8].Value.ToString();
                txtSales.Text = dataGridView5.CurrentRow.Cells[9].Value.ToString();
                txtWarranty.Text = dataGridView5.CurrentRow.Cells[10].Value.ToString();
                txtClient.Text = dataGridView5.CurrentRow.Cells[11].Value.ToString();
                txtAccessories.Text = dataGridView5.CurrentRow.Cells[12].Value.ToString();
                txtVer.Text = dataGridView5.CurrentRow.Cells[13].Value.ToString();
                txtReturnCause.Text = dataGridView5.CurrentRow.Cells[14].Value.ToString();
                txtFinishMark.Text = "";
                // txtBulidDate = dataGridView5.CurrentRow.Cells[15].Value.ToString();
                // dataGridView1.Rows[dataGridView5.CurrentRow.Index].DefaultCellStyle.BackColor = Color.Yellow;
                TBBuildDay.Text = dataGridView5.CurrentRow.Cells[15].Value.ToString();//建單日
                TBFixedDay.Text = dataGridView5.CurrentRow.Cells[18].Value.ToString();//維修完成日txtRepairFinD
                TBTestedDay.Text = dataGridView5.CurrentRow.Cells[20].Value.ToString();//測試完成日txtCalFinD
                TBFQCFinish.Text = dataGridView5.CurrentRow.Cells[22].Value.ToString();//FQC完成日txtFQCFinD
                TBCloseDay.Text = dataGridView5.CurrentRow.Cells[24].Value.ToString();//結案完成日txtCaseCloseD
                txtOldFixed.Text = dataGridView5.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                txtOldFixed2.Text = dataGridView5.CurrentRow.Cells[17].Value.ToString();//維修內容txtFixed
                                                                                        //2020/2/8  added
                if (dataGridView5.CurrentRow.Cells[28].Value.ToString() == "" || dataGridView5.CurrentRow.Cells[28].Value.ToString() == " ")
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = " ";// dateTimePicker1.Text = NULL;
                }
                else
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Long;
                    dateTimePicker1.Text = dataGridView5.CurrentRow.Cells[28].Value.ToString();//出貨日dateTimepick 
                }
                textVolt.Text = dataGridView5.CurrentRow.Cells[29].Value.ToString();//使用電壓texeVolt                    //2020/2/8  added
                CalDays();
            }
            bookID = Convert.ToInt32(dataGridView5.CurrentRow.Cells[0].Value.ToString());
        }
        private void dgvBook_DoubleClick6(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            btnDelete.Enabled = false;
            Clear();
            if (DGVFactoryRepair.CurrentRow.Index != -1)
            {
                DGVFactoryRepair.Rows[DGVFactoryRepair.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                DGVFactoryRepair.Rows[DGVFactoryRepair.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                txtStatus.Text = DGVFactoryRepair.CurrentRow.Cells[1].Value.ToString();
                //txtRepairOrCal.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                txtFinishOrSemi.Text = DGVFactoryRepair.CurrentRow.Cells[4].Value.ToString();//4
                txtModelName.Text = DGVFactoryRepair.CurrentRow.Cells[5].Value.ToString();//5
                txtBoardName.Text = DGVFactoryRepair.CurrentRow.Cells[6].Value.ToString();//6
                txtFinishSN.Text = DGVFactoryRepair.CurrentRow.Cells[7].Value.ToString();//7
                //txtSemiSN.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                txtBranch.Text = DGVFactoryRepair.CurrentRow.Cells[8].Value.ToString();//8
                txtSales.Text = DGVFactoryRepair.CurrentRow.Cells[11].Value.ToString();//11
                txtReturnCause.Text = DGVFactoryRepair.CurrentRow.Cells[9].Value.ToString();//9
                txtFinishMark.Text = "";
                //  txtClient.Text = "";
                // DGVFactoryRepair.Rows[DGVFactoryRepair.CurrentRow.Index].DefaultCellStyle.BackColor = Color.Yellow;
                bookID = Convert.ToInt32(DGVFactoryRepair.CurrentRow.Cells[0].Value.ToString());
            }
        }
        private void btnSearch_Click_1(object sender, EventArgs e)
        {
            //string cs = @"server=localhost;userid=root;password=1010;database=rma";
            string cs = @"server=192.168.1.31;port=36288;userid=rma;password=GdUmm0J4EnJZneue;database=rma;charset=utf8";
            string[] space = new string[30];
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add(new DataColumn("流水號", typeof(string)));
            dt.Columns.Add(new DataColumn("狀態", typeof(string)));
            dt.Columns.Add(new DataColumn("維修校驗", typeof(string)));
            dt.Columns.Add(new DataColumn("成品半成品", typeof(string)));
            dt.Columns.Add(new DataColumn("機種名", typeof(string)));
            dt.Columns.Add(new DataColumn("板名", typeof(string)));
            dt.Columns.Add(new DataColumn("成品序號", typeof(string)));
            dt.Columns.Add(new DataColumn("半成品序號", typeof(string)));
            dt.Columns.Add(new DataColumn("送件據點", typeof(string)));
            dt.Columns.Add(new DataColumn("所屬業務", typeof(string)));
            dt.Columns.Add(new DataColumn("保固內", typeof(string)));
            dt.Columns.Add(new DataColumn("客戶名", typeof(string)));
            dt.Columns.Add(new DataColumn("配件", typeof(string)));
            dt.Columns.Add(new DataColumn("版本", typeof(string)));
            dt.Columns.Add(new DataColumn("故障描述", typeof(string)));
            dt.Columns.Add(new DataColumn("建單日", typeof(string)));
            dt.Columns.Add(new DataColumn("建單人", typeof(string)));
            dt.Columns.Add(new DataColumn("維修內容", typeof(string)));
            dt.Columns.Add(new DataColumn("維修完成日", typeof(string)));
            dt.Columns.Add(new DataColumn("維修人員", typeof(string)));
            dt.Columns.Add(new DataColumn("測試完成日", typeof(string)));
            dt.Columns.Add(new DataColumn("測試人員", typeof(string)));
            dt.Columns.Add(new DataColumn("FQC完成日", typeof(string)));
            dt.Columns.Add(new DataColumn("FQC人員", typeof(string)));
            dt.Columns.Add(new DataColumn("結案完成日", typeof(string)));
            dt.Columns.Add(new DataColumn("結案人員", typeof(string)));
            dt.Columns.Add(new DataColumn("結案描述", typeof(string)));
            dt.Columns.Add(new DataColumn("測試故障描述", typeof(string)));
            dt.Columns.Add(new DataColumn("出貨日", typeof(string)));    //2020/2/8  added
            dt.Columns.Add(new DataColumn("使用電壓", typeof(string)));   //2020/2/8  added
            MySqlConnection con = new MySqlConnection(cs);
            try
            {
                con.Open();//開啟通道，建立連線，可能出現異常,使用try catch語句
                string sql = $"SELECT * FROM rma.rmarawdata WHERE 維修校驗 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 成品半成品 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 流水號 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 狀態 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 機種名 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 板名 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 成品序號 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 送件據點 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 所屬業務 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 保固內 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 客戶名 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 成品半成品 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 配件 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 版本 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 故障描述 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 建單日 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 建單人 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 維修內容 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 維修完成日 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 維修人員 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 測試完成日 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 測試人員 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| FQC完成日 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| FQC人員 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 結案完成日 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 結案人員 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 測試故障描述 LIKE CONCAT('%" + txtSearch.Text + "%')" +
                  "|| 結案描述 LIKE CONCAT('%" + txtSearch.Text + "%')";
                using var cmd = new MySqlCommand(sql, con);
                using MySqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        space[i] = rdr[i].ToString();
                    }
                    dt.Rows.Add(space);
                }
                dataGridView5.DataSource = dt;
                dataGridView5.DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView5.DefaultCellStyle.BackColor = Color.Beige;
                dataGridView5.Columns[0].Visible = false;
                con.Close();
                tabControl1.SelectedTab = tabPage5;
                btnOutput.Enabled = true;
            }
            catch (MySqlException)
            {
                MessageBox.Show("SQL SERVER連線異常!!!");
            }
            finally
            {
                con.Close();
            }
        }

        private void btnCancel_Click_1(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
            btnSave.Text = "輸入";
            Clear();
            GridFill();
        }
        private void btnDelete_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("確定刪除此資料?", "確認訊息", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.OK)
            {
                using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
                {
                    mysqlCon.Open();
                    MySqlCommand mySqlCmd = new MySqlCommand("RMADeleteByID", mysqlCon);
                    mySqlCmd.CommandType = CommandType.StoredProcedure;
                    mySqlCmd.Parameters.AddWithValue("_BookID", bookID);
                    mySqlCmd.ExecuteNonQuery();
                    MessageBox.Show("成功刪除資料");
                    Clear();
                    GridFill();
                }
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)///Login
        {
            dataGridView1.Enabled = false;
            dataGridView2.Enabled = false;
            dataGridView3.Enabled = false;
            dataGridView4.Enabled = false;
            groupBoxRMA.Enabled = false;
            groupBoxCom.Enabled = false;
            txtFixed.Enabled = false;
            textFactoryFixed.Enabled = false;
            btnRepairFinish.Enabled = false;
            BtnFactoryFixed.Enabled = false;
            btnDoNotRepair.Enabled = false;
            btnDoNotRepair.Enabled = false;
            btnTestFinish.Enabled = false;
            btnTestSentRepair.Enabled = false;
            btnFQCFinish.Enabled = false;
            btnFQCSentTest.Enabled = false;
            btnSave.Enabled = false;
            btnComSave.Enabled = false;
            btnDelete.Enabled = false;
            btnComUpload.Enabled = false;
            // string cs1 = @"server=localhost;userid=root;password=1010;database=rma";
            string cs1 = @"server=192.168.1.31;port=36288;userid=rma;password=GdUmm0J4EnJZneue;database=rma;charset=utf8";
            string[] space1 = new string[5];
            bool loginSuccess = false;
            System.Data.DataTable dtn = new System.Data.DataTable();
            dtn.Columns.Add(new DataColumn("流水號", typeof(string)));
            dtn.Columns.Add(new DataColumn("Name", typeof(string)));
            dtn.Columns.Add(new DataColumn("Password", typeof(string)));
            dtn.Columns.Add(new DataColumn("Show", typeof(string)));
            dtn.Columns.Add(new DataColumn("Role", typeof(string)));
            MySqlConnection con = new MySqlConnection(cs1);
            try
            {
                con.Open();//開啟通道，建立連線，可能出現異常,使用try catch語句
                string sql = "SELECT * FROM rma.rmauser;";
                using var cmd = new MySqlCommand(sql, con);
                using MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    // MessageBox.Show(rdr.FieldCount.ToString());
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        // space = rdr.FieldCount.ToString();
                        space1[i] = rdr[i].ToString();
                    }
                    dtn.Rows.Add(space1);
                }
                for (int i = 0; i < dtn.Rows.Count; i++)
                {
                    if (dtn.Rows[i][1].ToString() == txtUserID.Text)//查詢條件
                    {
                        if (dtn.Rows[i][2].ToString() == txtPassword.Text)//password
                        {
                            txtBulidPerson = dtn.Rows[i][3].ToString();//SHOW name
                            txtUserNameShow.Text = txtBulidPerson;
                            loginSuccess = true;
                            // 儲存使用者名稱和密碼到設定值
                            Properties.Settings.Default.UserName = txtUserID.Text;
                            Properties.Settings.Default.Password = txtPassword.Text;
                            Properties.Settings.Default.Save();

                            //  MessageBox.Show("登入為" + txtBulidPerson);
                            switch (dtn.Rows[i][4].ToString())
                            {
                                case "sales":
                                    txtUserNameShow.Enabled = false;
                                    txtUserNameShow.Items.Clear();
                                    txtUserNameShow.BackColor = Color.AntiqueWhite;
                                    dataGridView1.Enabled = true;
                                    dataGridView2.Enabled = true;
                                    dataGridView3.Enabled = true;
                                    dataGridView4.Enabled = true;
                                    txtFixed.Enabled = false;
                                    textFactoryFixed.Enabled = false;
                                    btnRepairFinish.Enabled = false;
                                    BtnFactoryFixed.Enabled = false;
                                    btnDoNotRepair.Enabled = false;
                                    btnDoNotRepair.Enabled = false;
                                    btnTestFinish.Enabled = false;
                                    btnTestSentRepair.Enabled = false;
                                    btnFQCFinish.Enabled = false;
                                    btnFQCSentTest.Enabled = false;
                                    btnDelete.Enabled = false;
                                    txtBulidPerson = dtn.Rows[i][3].ToString();//SHOW name
                                    txtUserNameShow.Text = txtBulidPerson;
                                    MessageBox.Show("登入為" + txtBulidPerson);
                                    btnSave.Enabled = true;
                                    btnComSave.Enabled = true;
                                    groupBoxRMA.Enabled = true;
                                    groupBoxCom.Enabled = true;
                                    ///
                                    DGVcom.Enabled = true;
                                    txtComNumber.Enabled = true;
                                    txtComModel.Enabled = true;
                                    dateTimePickerCom.Enabled = true;
                                    txtComCustomer.Enabled = true;
                                    CBComWarranty.Enabled = true;
                                    txtComAppearance.Enabled = true;
                                    CBComAppearanceSort.Enabled = true;
                                    ///
                                    txtComCause.Enabled = false;
                                    CBComCasueSort.Enabled = false;
                                    txtImprovement.Enabled = false;
                                    txtComPerson.Enabled = false;
                                    txtComImNow.Enabled = false;
                                    CBComCur.Enabled = false;
                                    CBComDepartment.Enabled = false;
                                    dateTimePickerComFinish.Enabled = false;
                                    ///
                                    PlayRole = "sales";
                                    break;
                                case "eng":
                                    txtUserNameShow.Enabled = true;
                                    txtUserNameShow.Items.Clear();
                                    txtUserNameShow.BackColor = Color.Orange;
                                    txtUserNameShow.Enabled = true;
                                    dataGridView1.Enabled = true;
                                    dataGridView2.Enabled = false;
                                    dataGridView3.Enabled = false;
                                    dataGridView4.Enabled = false;
                                    btnRepairFinish.Enabled = true;
                                    BtnFactoryFixed.Enabled = true;
                                    btnDoNotRepair.Enabled = true;
                                    btnTestFinish.Enabled = false;
                                    btnTestSentRepair.Enabled = false;
                                    btnFQCFinish.Enabled = false;
                                    btnFQCSentTest.Enabled = false;
                                    btnDelete.Enabled = false;
                                    txtRepairPerson = dtn.Rows[i][3].ToString();//SHOW name
                                    txtUserNameShow.Items.Add("蘇志鵬");
                                    txtUserNameShow.Items.Add("李錫坤");
                                    txtUserNameShow.Items.Add("藍文常");
                                    txtUserNameShow.Items.Add("柯宇星");
                                    txtUserNameShow.Items.Add("賴彥旭");
                                    txtUserNameShow.Text = txtRepairPerson;
                                    MessageBox.Show("登入為" + txtRepairPerson);
                                    btnSave.Enabled = false;
                                    btnComSave.Enabled = false;
                                    txtFixed.Enabled = true;
                                    textFactoryFixed.Enabled = true;
                                    groupBoxRMA.Enabled = false;
                                    groupBoxCom.Enabled = false;
                                    PlayRole = "eng";
                                    //                              warning = true;//支援
                                    //                               timer1.Enabled = true;
                                    break;
                                case "test":
                                    txtUserNameShow.Enabled = false;
                                    txtUserNameShow.Items.Clear();
                                    txtUserNameShow.BackColor = Color.Khaki;
                                    dataGridView1.Enabled = false;
                                    dataGridView2.Enabled = true;
                                    dataGridView3.Enabled = false;
                                    dataGridView4.Enabled = false;
                                    txtFixed.Enabled = false;
                                    textFactoryFixed.Enabled = false;
                                    btnRepairFinish.Enabled = false;
                                    BtnFactoryFixed.Enabled = false;
                                    btnDoNotRepair.Enabled = false;
                                    btnTestFinish.Enabled = true;
                                    btnTestSentRepair.Enabled = true;
                                    btnFQCFinish.Enabled = false;
                                    btnFQCSentTest.Enabled = false;
                                    btnDelete.Enabled = false;
                                    txtCalPerson = dtn.Rows[i][3].ToString();//SHOW name
                                    txtUserNameShow.Text = txtCalPerson;
                                    MessageBox.Show("登入為" + txtCalPerson);
                                    btnSave.Enabled = false;
                                    btnComSave.Enabled = false;
                                    groupBoxRMA.Enabled = false;
                                    groupBoxCom.Enabled = false;

                                    PlayRole = "test";
                                    break;
                                case "FQC":
                                    txtUserNameShow.Enabled = false;
                                    txtUserNameShow.Items.Clear();
                                    txtUserNameShow.BackColor = Color.PaleGreen;
                                    dataGridView1.Enabled = false;
                                    dataGridView2.Enabled = false;
                                    dataGridView3.Enabled = true;
                                    dataGridView4.Enabled = false;
                                    txtFixed.Enabled = false;
                                    textFactoryFixed.Enabled = false;
                                    btnRepairFinish.Enabled = false;
                                    BtnFactoryFixed.Enabled = false;
                                    btnDoNotRepair.Enabled = false;
                                    btnTestFinish.Enabled = false;
                                    btnTestSentRepair.Enabled = false;
                                    btnFQCFinish.Enabled = true;
                                    btnFQCSentTest.Enabled = true;
                                    btnDelete.Enabled = false;
                                    txtFQCPerson = dtn.Rows[i][3].ToString();//SHOW name
                                    txtUserNameShow.Text = txtFQCPerson;
                                    MessageBox.Show("登入為" + txtFQCPerson);
                                    btnSave.Enabled = false;
                                    //btnComSave.Enabled = true;
                                    groupBoxRMA.Enabled = false;
                                    groupBoxCom.Enabled = true;
                                    ///
                                    DGVcom.Enabled = true;
                                    txtComCause.Enabled = true;
                                    CBComCasueSort.Enabled = true;
                                    txtImprovement.Enabled = true;
                                    txtComPerson.Enabled = true;
                                    CBComDepartment.Enabled = true;
                                    dateTimePickerComFinish.Enabled = true;
                                    txtComImNow.Enabled = true;
                                    CBComCur.Enabled = true;
                                    ///
                                    txtComNumber.Enabled = false;
                                    txtComModel.Enabled = false;
                                    dateTimePickerCom.Enabled = false;
                                    txtComCustomer.Enabled = false;
                                    CBComWarranty.Enabled = false;
                                    txtComAppearance.Enabled = false;
                                    CBComAppearanceSort.Enabled = false;
                                    ///
                                    PlayRole = "FQC";
                                    break;
                                case "MA":
                                    BTNDeadLine.Enabled = true;
                                    PlayRole = "MA";
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
                if (loginSuccess == false)
                {
                    MessageBox.Show("無此使用者!!");
                    loginSuccess = false;
                }
                con.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("save SERVER連線異常!!!");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                txtUserID.Text = "";
                txtPassword.Text = "";

                con.Close();
            }
        }

        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }

        private void dataGridView2_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                dataGridView2.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }
        private void dataGridView3_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                dataGridView3.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                dataGridView3.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }
        private void dataGridView4_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                dataGridView4.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                dataGridView4.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }
        private void dataGridView5_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                dataGridView5.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                dataGridView5.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }
        private void DGVFactoryRepair_CellLeave_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                DGVFactoryRepair.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                DGVFactoryRepair.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }
        private void DGVFactoryFixed_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                DGVFactoryFixed.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                DGVFactoryFixed.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }

        private void btnTestSentRepair_Click(object sender, EventArgs e)//送回維修
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            conn.Open();//開啟通道，建立連線，出現異常時,使用try catch語句
            using var cmd = new MySqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + "待維修" +
                    "' WHERE 流水號='" + bookID + "'";
            cmd.ExecuteNonQuery();
            conn.Close();
            MessageBox.Show("成功更新資料!");
            Clear();
            GridFill();
        }

        private void btnFQCSentTest_Click(object sender, EventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            conn.Open();//開啟通道，建立連線，出現異常時,使用try catch語句
            using var cmd = new MySqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + "待測試" +
                    "' WHERE 流水號='" + bookID + "'";
            cmd.ExecuteNonQuery();
            conn.Close();
            MessageBox.Show("成功更新資料!");
            Clear();
            GridFill();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridFill();
        }
        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridFill();
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            ////////////////////////////////////////////////////////////////////////
            DataTable dt = new DataTable();
            //dt.Columns.Add("狀態");
            //dt.Columns.Add("維修校驗");
            //dt.Columns.Add("成品半成品");
            //dt.Columns.Add("機種名");
            //dt.Columns.Add("板名");
            foreach (DataGridViewColumn column in dataGridView5.Columns)
                dt.Columns.Add(column.Name);
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                dt.Rows.Add();
                for (int j = 0; j < dataGridView5.Columns.Count; j++)
                {
                    dt.Rows[i][j] = dataGridView5.Rows[i].Cells[j].Value.ToString();
                }
            }
            MemoryStream stream = new MemoryStream();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            // 關閉新許可模式通知
            using (var ep1 = new ExcelPackage(stream))
            {
                var worksheet = ep1.Workbook.Worksheets.Add("RMA匯出");
                // ExcelWorksheet sheet = ep1.Workbook.Worksheets[0];
                //datatable 使用 LoadFromDatatable,collection 可使用 LoadFromCollection
                worksheet.Cells["A1"].LoadFromDataTable(dt, true);
                ep1.Save();
                string fileName = "RMA" + DateTime.Now.ToString("yyyyMMddHHmm");
                // string path1 = AppDomain.CurrentDomain.BaseDirectory + "xlsx\\"+fileName+".xlsx";
                // string filePath1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\xlsx\\setLotFile.xlsx";
                string path1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + fileName + ".xlsx";
                using FileStream myFile = File.Open(path1, FileMode.OpenOrCreate);
                stream.WriteTo(myFile);
                myFile.Close();
                MessageBox.Show("已成功存檔於桌面，檔名:" + fileName + ".xlsx");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //实例化打印对象
            PrintDocument printDocument1 = new PrintDocument();
            //////设置打印用的纸张,当设置为Custom的时候，可以自定义纸张的大小
            printDocument1.DefaultPageSettings.PaperSize = new PaperSize("Custum", 800, 550);
            //注册PrintPage事件，打印每一页时会触发该事件
            printDocument1.PrintPage += new PrintPageEventHandler(this.PrintDocument_PrintPage);
            //初始化打印预览对话框对象
            PrintPreviewDialog printPreviewDialog1 = new PrintPreviewDialog();
            //将printDocument1对象赋值给打印预览对话框的Document属性
            printPreviewDialog1.Document = printDocument1;
            //打开打印预览对话框
            DialogResult result = printPreviewDialog1.ShowDialog();
            if (result == DialogResult.OK)
                printDocument1.Print();//开始打印
        }
        private void PrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            txtBulidDate = DateTime.Now.ToLongDateString().ToString();//建單日時
                                                                      //设置打印内容及其字体，颜色和位置
            e.Graphics.DrawString("-----------------------------------------------------------------------------------", new Font(new FontFamily("新細明體"), 24), System.Drawing.Brushes.Black, 10, 0);
            e.Graphics.DrawString("客戶【" + txtRepairOrCal.Text + "】單", new Font(new FontFamily("新細明體"), 24), System.Drawing.Brushes.Black, 300, 25);
            e.Graphics.DrawString("客戶: 【" + txtClient.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 50, 75);
            e.Graphics.DrawString("建單人員: 【" + txtUserNameShow.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 300, 75);
            e.Graphics.DrawString("建單日期: " + txtBulidDate, new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 550, 75);
            //e.Graphics.DrawString("客戶:" + txtClient.Text, new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 50, 75);   
            e.Graphics.DrawString("機種(型號): 【" + txtModelName.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 50, 100);
            e.Graphics.DrawString("成品序號: 【" + txtFinishSN.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 300, 100);
            e.Graphics.DrawString("使用電壓: 【" + textVolt.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 550, 100);   ///new 2022/2/8
            e.Graphics.DrawString("半成品板號: 【" + txtBoardName.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 50, 125);
            e.Graphics.DrawString("半成品序號: 【" + txtSemiSN.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 300, 125);
            e.Graphics.DrawString("出貨日: 【" + dateTimePicker1.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 550, 125); // new 2022/2/8
            e.Graphics.DrawString("配件: 【" + txtAccessories.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 50, 150);
            e.Graphics.DrawString("版本: 【" + txtVer.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 300, 150);
            e.Graphics.DrawString("保固內: 【" + txtWarranty.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 550, 150);
            e.Graphics.DrawString("送回原因: 【" + txtReturnCause.Text + "】", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 100, 200);
            e.Graphics.DrawString("處理內容:_________________________________________________________________________________________", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 100, 300);
            e.Graphics.DrawString("         _________________________________________________________________________________________", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 100, 325);
            e.Graphics.DrawString("         _________________________________________________________________________________________", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 100, 350);
            e.Graphics.DrawString("測試人員簽名: _____________   維修人員簽名: _____________    品管人員簽名: _____________", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 10, 400);
            e.Graphics.DrawString("白聯:業務留存                      黃聯:品管留存                            紅聯:工廠留存", new Font(new FontFamily("新細明體"), 14), System.Drawing.Brushes.Black, 50, 425);
            e.Graphics.DrawString("-----------------------------------------------------------------------------------", new Font(new FontFamily("新細明體"), 24), System.Drawing.Brushes.Black, 10, 450);
        }

        private void btnDoNotRepair_Click(object sender, EventArgs e)
        {
            txtBulidDate = DateTime.Now.ToLocalTime().ToString();//建單日時
            ////SQL  開始write //////////////////////////////////////////////////////////////////////////////////////////////    
            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();//開啟通道，建立連線，出現異常時,使用try catch語句
                using var cmd = new MySqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = "UPDATE rma.rmarawdata SET 狀態='" + "待結案" +
                            "',維修內容='" + txtOldFixed.Text + "#" + txtFixed.Text +
                            "',維修完成日='" + DateTime.Now.ToLocalTime().ToString() +
                            "',維修人員 ='" + txtRepairPerson +
                             "' WHERE 流水號='" + bookID + "'";
                cmd.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("成功更新資料!");
                Clear();
                GridFill();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Save SERVER連線異常!!!");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void txtUserNameShow_SelectionChangeCommitted(object sender, EventArgs e)//ENG 下拉改變時更新名字
        {
            txtRepairPerson = txtUserNameShow.SelectedItem.ToString();
        }

        private void BtnFactoryFixed_Click(object sender, EventArgs e)
        {
            txtBulidDate = DateTime.Now.ToLocalTime().ToString();//建單日時
            MySqlConnection con = new MySqlConnection(connectionString);
            try
            {
                con.Open();//開啟通道，建立連線，可能出現異常,使用try catch語句
                using var cmd = new MySqlCommand();
                cmd.Connection = con;
                //將先前記錄工站修改為ex//2021   
                cmd.CommandText = "UPDATE rma.micrawdata SET 工站 = '組裝-EX' where(序號 ='" + txtFinishSN.Text + "' and 工站 = '組裝')LIMIT 5;";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "UPDATE rma.micrawdata SET 工站 = '測試-EX' where(序號 ='" + txtFinishSN.Text + "' and 工站 = '測試')LIMIT 5;";
                cmd.ExecuteNonQuery();
                // cmd.CommandText = "UPDATE rma.micrawdata SET 工站 = '內檢-EX' where(序號 ='" + txtFinishSN.Text + "' and 工站 = '內檢')LIMIT 5;";
                //  cmd.ExecuteNonQuery();
                cmd.CommandText = "UPDATE rma.micrawdata SET 工站 = 'FQC-EX' where(序號 ='" + txtFinishSN.Text + "' and 工站 = 'FQC')LIMIT 5;";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "UPDATE rma.micrawdata SET 工站 = '包裝-EX' where(序號 ='" + txtFinishSN.Text + "' and 工站 = '包裝')LIMIT 5;";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "UPDATE rma.micrawdata SET 工站 = '送修-EX' where(序號 ='" + txtFinishSN.Text + "' and 工站 = '送修')LIMIT 5;";
                cmd.ExecuteNonQuery();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("rma.micrawdata SQL SERVER連線異常!!!");
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("無法連線至Server， 請連絡資訊人員");
                        break;
                    case 1045:
                        MessageBox.Show("登入帳號或密碼錯誤，請連絡資訊人員");
                        break;
                    default:
                        break;
                }
            }
            finally
            {
                con.Close();
            }
            ////SQL  開始write //////////////////////////////////////////////////////////////////////////////////////////////    
            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();//開啟通道，建立連線，出現異常時,使用try catch語句
                using var cmd = new MySqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = "UPDATE rma.micfactoryrepair SET 狀態='" + "維修完成" +
                            "',維修內容='" + textFactoryFixed.Text +
                            "',維修完成日='" + DateTime.Now.ToLocalTime().ToString() +
                            "',維修人員 ='" + txtRepairPerson +
                             "' WHERE 流水號='" + bookID + "'";
                cmd.ExecuteNonQuery();
                conn.Close();

            }
            catch (MySqlException ex)
            {
                MessageBox.Show("save SERVER連線異常!!!");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close();
            }
            MessageBox.Show("成功更新資料!");
            Clear();
            GridFill();
        }

        private void DGVFactoryFixed_Click(object sender, EventArgs e)
        {
            Clear();
            if (DGVFactoryFixed.CurrentRow.Index != -1)
            {
                DGVFactoryFixed.Rows[DGVFactoryFixed.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                DGVFactoryFixed.Rows[DGVFactoryFixed.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                btnSave.Enabled = false;
                btnDelete.Enabled = false;
                txtStatus.Text = DGVFactoryFixed.CurrentRow.Cells[1].Value.ToString();
                //txtRepairOrCal.Text = DGVFactoryFixed.CurrentRow.Cells[2].Value.ToString();
                txtFinishOrSemi.Text = DGVFactoryFixed.CurrentRow.Cells[4].Value.ToString();
                txtModelName.Text = DGVFactoryFixed.CurrentRow.Cells[5].Value.ToString();
                txtBoardName.Text = DGVFactoryFixed.CurrentRow.Cells[6].Value.ToString();
                txtFinishSN.Text = DGVFactoryFixed.CurrentRow.Cells[7].Value.ToString();
                //txtSemiSN.Text = DGVFactoryFixed.CurrentRow.Cells[7].Value.ToString();
                txtBranch.Text = DGVFactoryFixed.CurrentRow.Cells[8].Value.ToString();
                txtSales.Text = DGVFactoryFixed.CurrentRow.Cells[11].Value.ToString();
                txtReturnCause.Text = DGVFactoryFixed.CurrentRow.Cells[9].Value.ToString();
                txtFinishMark.Text = "";
                // DGVFactoryFixed.Rows[DGVFactoryFixed.CurrentRow.Index].DefaultCellStyle.BackColor = Color.Yellow;
                bookID = Convert.ToInt32(DGVFactoryFixed.CurrentRow.Cells[0].Value.ToString());
            }
        }

        private void btnFactorySearch_Click(object sender, EventArgs e)
        {
            string cs = @"server=192.168.1.31;port=36288;userid=rma;password=GdUmm0J4EnJZneue;database=rma;charset=utf8";
            string[] space = new string[16];
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add(new DataColumn("流水號", typeof(string)));
            dt.Columns.Add(new DataColumn("狀態", typeof(string)));
            dt.Columns.Add(new DataColumn("需求日期", typeof(string)));
            dt.Columns.Add(new DataColumn("工令", typeof(string)));
            dt.Columns.Add(new DataColumn("成品半成品", typeof(string)));
            dt.Columns.Add(new DataColumn("型號", typeof(string)));
            dt.Columns.Add(new DataColumn("板名", typeof(string)));
            dt.Columns.Add(new DataColumn("序號", typeof(string)));
            dt.Columns.Add(new DataColumn("送修單位", typeof(string)));
            dt.Columns.Add(new DataColumn("故障描述", typeof(string)));
            dt.Columns.Add(new DataColumn("建單日", typeof(string)));
            dt.Columns.Add(new DataColumn("建單人", typeof(string)));
            dt.Columns.Add(new DataColumn("維修內容", typeof(string)));
            dt.Columns.Add(new DataColumn("維修完成日", typeof(string)));
            dt.Columns.Add(new DataColumn("維修人員", typeof(string)));
            dt.Columns.Add(new DataColumn("結案描述", typeof(string)));

            MySqlConnection con = new MySqlConnection(cs);
            try
            {
                con.Open();//開啟通道，建立連線，可能出現異常,使用try catch語句
                string sql = $"SELECT * FROM rma.micfactoryrepair WHERE 流水號 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 狀態 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 工令 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 成品半成品 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 型號 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 板名 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 序號 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 送修單位 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 故障描述 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 建單日 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 建單人 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 維修內容 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 維修完成日 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 維修人員 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')" +
                  "|| 結案描述 LIKE CONCAT('%" + txtFactoryQuery.Text + "%')";

                using var cmd = new MySqlCommand(sql, con);
                using MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        space[i] = rdr[i].ToString();
                    }
                    dt.Rows.Add(space);
                }
                dataGridViewFactoryQuery.DataSource = dt;
                dataGridViewFactoryQuery.DefaultCellStyle.ForeColor = Color.Blue;
                dataGridViewFactoryQuery.DefaultCellStyle.BackColor = Color.Beige;
                dataGridViewFactoryQuery.Columns[0].Visible = false;
                con.Close();
                tabControl3.SelectedTab = tabPage10;
                btnFactoryOutput.Enabled = true;
            }
            catch (MySqlException)
            {
                MessageBox.Show("SQL SERVER連線異常!!!");
            }
            finally
            {
                con.Close();
            }
        }

        private void btnFactoryOutput_Click(object sender, EventArgs e)
        {
            ////////////////////////////////////////////////////////////////////////
            DataTable dt = new DataTable();
            foreach (DataGridViewColumn column in dataGridViewFactoryQuery.Columns)
                dt.Columns.Add(column.Name);
            for (int i = 0; i < dataGridViewFactoryQuery.Rows.Count; i++)
            {
                dt.Rows.Add();
                for (int j = 0; j < dataGridViewFactoryQuery.Columns.Count; j++)
                {
                    dt.Rows[i][j] = dataGridViewFactoryQuery.Rows[i].Cells[j].Value.ToString();
                }
            }
            MemoryStream stream = new MemoryStream();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            // 關閉新許可模式通知
            using (var ep1 = new ExcelPackage(stream))
            {
                var worksheet = ep1.Workbook.Worksheets.Add("廠內維修匯出");
                // ExcelWorksheet sheet = ep1.Workbook.Worksheets[0];
                //datatable 使用 LoadFromDatatable,collection 可使用 LoadFromCollection
                worksheet.Cells["A1"].LoadFromDataTable(dt, true);
                ep1.Save();
                string fileName = "廠內維修" + DateTime.Now.ToString("yyyyMMddHHmm");
                // string path1 = AppDomain.CurrentDomain.BaseDirectory + "xlsx\\"+fileName+".xlsx";
                // string filePath1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\xlsx\\setLotFile.xlsx";
                string path1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + fileName + ".xlsx";
                using FileStream myFile = File.Open(path1, FileMode.OpenOrCreate);
                stream.WriteTo(myFile);
                myFile.Close();
                MessageBox.Show("已成功存檔於桌面，檔名:" + fileName + ".xlsx");
            }
        }

        private void btnItemsQuery_Click(object sender, EventArgs e)
        {

            //string cs = @"server=192.168.1.31;port=36288;userid=rma;password=GdUmm0J4EnJZneue;database=rma;charset=utf8";
            string connectString = "Data Source=192.168.1.25; Initial Catalog=MICROTEST; User ID=stalker; Password=2u/3vupw.djo5k3";
            System.Data.DataTable dt = new DataTable();
            SqlConnection sqlConnection = new SqlConnection(connectString);
            try
            {

                sqlConnection.Open();
                if (RBtn1.Checked == true)
                {
                    string[] space = new string[4];
                    //庫存
                    String sqlString = $@"select Top 500 INVMB.*,
                    A.MA007 as MB005C1,
                    A.MA008 as MB006C1,
                    A.MA009 as MB007C1,
                    A.MA010 as MB008C1, 
                    B.MA003 as MB005C,
                    C.MA003 as MB006C,
                    D.MA003 as MB007C,
                    E.MA003 as MB008C,
                    F.MC002 as MB017C,
                    (MB057 + MB058 + MB059 + MB060) as MB7890C,G.MD002 as MB068C,
                    H.MA002 as MB032C,I.MV002 as MB018C,
                    J.MV002 as MB067C
                    from MICROTEST..INVMB as INVMB WITH(NOLOCK)
                    left join MICROTEST..INVMA as B WITH(NOLOCK) on B.MA001 = '1' and MB005 = B.MA002
                    left join MICROTEST..INVMA as C WITH(NOLOCK) on C.MA001 = '2' and MB006 = C.MA002
                    left join MICROTEST..INVMA as D WITH(NOLOCK) on D.MA001 = '3' and MB007 = D.MA002
                    left join MICROTEST..INVMA as E WITH(NOLOCK) on E.MA001 = '4' and MB008 = E.MA002
                    left join MICROTEST..CMSMC as F WITH(NOLOCK) on F.MC001 = INVMB.MB017
                    left join MICROTEST..CMSMD as G WITH(NOLOCK) on G.MD001 = INVMB.MB068
                    left join MICROTEST..PURMA as H WITH(NOLOCK) on H.MA001 = INVMB.MB032
                    left join MICROTEST..CMSMV as I WITH(NOLOCK) on I.MV001 = INVMB.MB018
                    left join MICROTEST..CMSMV as J WITH(NOLOCK) on J.MV001 = INVMB.MB067
                    LEFT JOIN MICROTEST..CMSMA as A WITH(NOLOCK) ON '1' = '1' 
                    where(INVMB.MB001 like N'" + "%" + txtItemsQuery.Text + "%" + "') " +
                    "or(INVMB.MB002 like N'" + "%" + txtItemsQuery.Text + "%" + "') " +
                    "or(INVMB.MB003 like N'" + "%" + txtItemsQuery.Text + "%" + "') order by INVMB.MB001";
                    SqlCommand mySqlCmd = new SqlCommand(sqlString, sqlConnection);
                    //搜尋到的資料取出
                    //連結讀取資料庫資料的元件            執行ExecuteReader()
                    SqlDataReader dataReader = mySqlCmd.ExecuteReader();
                    dt.Columns.Add(new DataColumn("料號", typeof(string)));
                    dt.Columns.Add(new DataColumn("品名", typeof(string)));
                    dt.Columns.Add(new DataColumn("規格", typeof(string)));
                    dt.Columns.Add(new DataColumn("庫存", typeof(string)));

                    int rowCount = 0;
                    while (dataReader.Read())
                    {
                        rowCount++;
                        space[0] = dataReader["MB001"].ToString();
                        space[1] = dataReader["MB002"].ToString();
                        space[2] = dataReader["MB003"].ToString();
                        space[3] = dataReader["MB064"].ToString();
                        dt.Rows.Add(space);
                    }
                    labelCOUNT.Text = rowCount.ToString();
                    dataGridViewItemsQuery.DataSource = dt;
                    dataGridViewItemsQuery.DefaultCellStyle.ForeColor = Color.Blue;
                    dataGridViewItemsQuery.DefaultCellStyle.BackColor = Color.Beige;
                    //    dataGridViewItemsQuery.Columns[0].Visible = false;
                    dataReader.Close();
                }
                if (RBtn2.Checked == true)
                {
                    string[] space = new string[3];
                    //BOM
                    String sqlString = $@"select Top 500  BOMMC.*,
                A.MB002 as MC001C1, A.MB003 as MC001C2,A.MB025 as MC001C3,
                A.MB004 as MC001C4, A.MB072 as MC001C5,
                B.MQ002 as MC005C , C.TA003 as MC006C,
                F.MF002 AS MC018C
                from MICROTEST..BOMMC as BOMMC
                left join MICROTEST..INVMB as A ON A.MB001 = MC001
                left join MICROTEST..CMSMQ as B on B.MQ001 = MC005
                left join MICROTEST..BOMTA as C on C.TA001 = MC006 AND C.TA002 = MC007
                left join MICROTEST..ADMMF AS F ON F.MF001 = MC018
                where  ( BOMMC.MC001 like N'" + "%" + txtItemsQuery.Text + "%" + "')  " +
                //      "or( BOMMC.MC001C1 like N'" + "%" + txtItemsQuery.Text + "%" + "') " +
                //      "or( BOMMC.MC001C2 like N'" + "%" + txtItemsQuery.Text + "%" + "') " +
                "order by BOMMC.MC001";
                    SqlCommand mySqlCmd = new SqlCommand(sqlString, sqlConnection);
                    //搜尋到的資料取出
                    //連結讀取資料庫資料的元件            執行ExecuteReader()
                    SqlDataReader dataReader = mySqlCmd.ExecuteReader();
                    dt.Columns.Add(new DataColumn("料號", typeof(string)));
                    dt.Columns.Add(new DataColumn("品名", typeof(string)));
                    dt.Columns.Add(new DataColumn("規格", typeof(string)));
                    //  dt.Columns.Add(new DataColumn("庫存", typeof(string)));
                    //for (int i = 0; i < 50; i++)
                    //{
                    //    // space[i]=dataReader.GetName(i);
                    //    dt.Columns.Add(new DataColumn(dataReader.GetName(i), typeof(string)));
                    //}
                    ////  dt.Rows.Add(space);
                    int rowCount = 0;
                    while (dataReader.Read())
                    {
                        rowCount++;
                        space[0] = dataReader["MC001"].ToString();
                        space[1] = dataReader["MC001C1"].ToString();
                        space[2] = dataReader["MC001C2"].ToString();
                        //  space[3] = dataReader["MC064"].ToString();
                        dt.Rows.Add(space);
                    }
                    labelCOUNT.Text = rowCount.ToString();
                    dataGridViewItemsQuery.DataSource = dt;
                    dataGridViewItemsQuery.DefaultCellStyle.ForeColor = Color.Blue;
                    dataGridViewItemsQuery.DefaultCellStyle.BackColor = Color.Beige;
                    //    dataGridViewItemsQuery.Columns[0].Visible = false;
                    dataReader.Close();
                }
                if (RBtn3.Checked == true)
                {
                    string[] space = new string[4];
                    //BOM 細查

                    String sqlString = $@"select BOMMD.* ,
                     B.MB002 as MD003C1, B.MB003 as MD003C2, B.MB025 as MD003C3,
                     B.MB004 as MD003C4, B.MB072 as MD003C5,
                     MW002 as MD009C,B.MB030 AS MB030C,B.MB031 AS MB031C
                     from MICROTEST..BOMMD as BOMMD
                     left join MICROTEST..INVMB as B ON B.MB001 = BOMMD.MD003
                     left join MICROTEST..CMSMW as CMSMW ON MW001 = MD009
                     where BOMMD.MD001 = '" + txtItemsQuery.Text + "' order by BOMMD.MD001,BOMMD.MD002";
                    SqlCommand mySqlCmd = new SqlCommand(sqlString, sqlConnection);
                    SqlDataReader dataReader = mySqlCmd.ExecuteReader();
                    dt.Columns.Add(new DataColumn("料號", typeof(string)));
                    dt.Columns.Add(new DataColumn("品名", typeof(string)));
                    dt.Columns.Add(new DataColumn("規格", typeof(string)));
                    dt.Columns.Add(new DataColumn("用量", typeof(string)));
                    int rowCount = 0;
                    while (dataReader.Read())
                    {
                        rowCount++;
                        space[0] = dataReader["MD003"].ToString();
                        space[1] = dataReader["MD003C1"].ToString();
                        space[2] = dataReader["MD003C2"].ToString();
                        space[3] = dataReader["MD006"].ToString();
                        dt.Rows.Add(space);
                    }
                    labelCOUNT.Text = rowCount.ToString();
                    dataGridViewItemsQuery.DataSource = dt;
                    dataGridViewItemsQuery.DefaultCellStyle.ForeColor = Color.Blue;
                    dataGridViewItemsQuery.DefaultCellStyle.BackColor = Color.Beige;
                    //    dataGridViewItemsQuery.Columns[0].Visible = false;
                    dataReader.Close();
                }
            }
            catch (MySqlException)
            {
                MessageBox.Show("SQL SERVER連線異常!!!");
            }
            finally
            {
                //  dataReader.Close();
            }
        }



        private void BTNDeadLine_Click(object sender, EventArgs e)
        {
            if (DGVFactoryRepair.CurrentRow.Index != -1)
            {
                string ModelName = DGVFactoryRepair.CurrentRow.Cells[4].Value.ToString();
                string BoardName = DGVFactoryRepair.CurrentRow.Cells[5].Value.ToString();
                string SerialNumber = DGVFactoryRepair.CurrentRow.Cells[6].Value.ToString();
                string ID = DGVFactoryRepair.CurrentRow.Cells[0].Value.ToString();

                //MessageBox.Show("送修建立\n");
                Form2 popupform2 = new Form2(ModelName, BoardName, SerialNumber, ID);//傳入按的工令
                DialogResult dialogresult = popupform2.ShowDialog();

                if (dialogresult == DialogResult.OK)
                {
                    // Console.WriteLine("You clicked OK");
                }
                else if (dialogresult == DialogResult.Cancel)
                {
                    // Console.WriteLine("You clicked either Cancel or X button in the top right corner");
                }
                popupform2.Dispose();
            }
        }


        private void timer1_Tick_1(object sender, EventArgs e)
        {
            System.Data.DataTable dtn = new System.Data.DataTable();
            dtn.Columns.Add(new DataColumn("ID", typeof(string)));
            dtn.Columns.Add(new DataColumn("FROM", typeof(string)));
            dtn.Columns.Add(new DataColumn("CALLDATE", typeof(string)));
            dtn.Columns.Add(new DataColumn("FINISHDATE", typeof(string)));
            dtn.Columns.Add(new DataColumn("CLOSE", typeof(string)));
            // MySqlConnection con = new MySqlConnection(cs1);
            MySqlConnection con = new MySqlConnection(connectionString);
            try
            {
                con.Open();//開啟通道，建立連線，可能出現異常,使用try catch語句
                string sql = "SELECT count(*) FROM rma.micHELP where finishdate IS NULL;";
                using var cmd = new MySqlCommand(sql, con);
                using MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        if (rdr[i].ToString() != "0")
                        {
                            //MessageBox.Show("CALL");
                            //MessageBox.Show("送修建立\n");
                            Form3 popupform2 = new Form3();//傳入按的工令
                            DialogResult dialogresult = popupform2.ShowDialog();

                            if (dialogresult == DialogResult.OK)
                            {
                                // Console.WriteLine("You clicked OK");
                            }
                            else if (dialogresult == DialogResult.Cancel)
                            {
                                // Console.WriteLine("You clicked either Cancel or X button in the top right corner");
                            }
                            popupform2.Dispose();
                        }
                    }
                }
                con.Close();
            }



            catch (MySqlException)
            {
                MessageBox.Show("SQL SERVER連線異常!!!");
            }
            finally
            {
                con.Close();
            }


            //     MessageBox.Show("TIMER");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 讀取使用者名稱和密碼從設定值
            txtUserID.Text = Properties.Settings.Default.UserName;
            txtPassword.Text = Properties.Settings.Default.Password;
        }

        private void tabFunc_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridFillCOM();//客訴頁面
        }
        void GridFillCOM()
        {
            string SQLWord = "";
            if (radioButton1.Checked == true) { SQLWord = "SELECT * FROM rma.CustomerComplaint where 流水號 > 0 "; }//處理中
            if (radioButton2.Checked == true) { SQLWord = "SELECT * FROM rma.CustomerComplaint where 目前進度='待確認'"; }
            if (radioButton3.Checked == true) { SQLWord = "SELECT * FROM rma.CustomerComplaint where 目前進度='暫停中'"; }
            if (radioButton4.Checked == true) { SQLWord = "SELECT * FROM rma.CustomerComplaint where 目前進度='已完成'"; }
            switch (CBComDateBetween.Text)
            {
                case "最近一個月":
                    SQLWord += " AND 建單日 >= DATE_SUB(NOW(), INTERVAL 1 MONTH)";
                    break;
                case "最近一季":
                    SQLWord += " AND 建單日 >= DATE_SUB(NOW(), INTERVAL 3 MONTH)";
                    break;
                case "最近一年":
                    SQLWord += " AND 建單日 >= DATE_SUB(NOW(), INTERVAL 12 MONTH)";
                    break;
                default:
                    break;
            }
            try
            {
                using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
                {
                    mysqlCon.Open();
                    MySqlDataAdapter sqlDa = new MySqlDataAdapter(SQLWord, mysqlCon);
                    System.Data.DataTable dtblBook = new System.Data.DataTable();
                    sqlDa.Fill(dtblBook);
                    DGVcom.DataSource = dtblBook;
                    DGVcom.DefaultCellStyle.ForeColor = Color.Blue;
                    DGVcom.DefaultCellStyle.BackColor = Color.Beige;
                    DGVcom.Columns[0].Visible = false;//流水號ID
                                                      //DGVcom.Columns[1].Visible = false;//

                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("SQL SERVER連線異常!!!");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //mysqlCon.Close();
            }
        }
        // 当单选按钮的选中状态改变时，会触发这个事件处理器
        void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = sender as RadioButton;

            if (radioButton == null)
            {
                MessageBox.Show("Error in radioButton_CheckedChanged");
                return;
            }

            if (radioButton.Checked)
            {
                GridFillCOM();//客訴頁面
            }
        }

        private void btnComCancel_Click(object sender, EventArgs e)
        {
            if (PlayRole == "sales")
            { btnComSave.Enabled = true; }
            else { btnComSave.Enabled = false; }
            btnComSave.Text = "輸入";
            ClearCOM();
            GridFillCOM();
        }
        private void CBComDateBetween_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridFillCOM();
        }
        void ClearCOM()
        {
            txtComNumber.Text = ""; txtComModel.Text = ""; dateTimePickerCom.Format = DateTimePickerFormat.Custom; dateTimePickerCom.CustomFormat = " ";
            txtComCustomer.Text = ""; CBComWarranty.Text = ""; txtComAppearance.Text = ""; CBComAppearanceSort.Text = "";
            txtComCause.Text = ""; CBComCasueSort.Text = ""; CBComWarranty.Text = ""; txtImprovement.Text = ""; txtComPerson.Text = ""; CBComDepartment.Text = "";
            dateTimePickerComFinish.Format = DateTimePickerFormat.Custom; dateTimePickerComFinish.CustomFormat = " ";
            txtComImNow.Text = "";
            CBComCur.Text = "";
            bookComID = 0;
            btnComUpload.Enabled = false;
            btnComSave.Text = "輸入";
            txtStatus.Text = "";
            txtUserNameShow.Enabled = false;
            txtUserNameShow.Items.Clear();
        }

        private void DGVcom_Click(object sender, EventArgs e)
        {
            btnComSave.Enabled = false;
            btnComUpload.Enabled = true;
            if (DGVcom.CurrentRow.Index != -1)
            {
                DGVcom.Rows[DGVcom.CurrentRow.Index].DefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
                DGVcom.Rows[DGVcom.CurrentRow.Index].DefaultCellStyle.ForeColor = Color.White;
                bookComID = Convert.ToInt32(DGVcom.CurrentRow.Cells[0].Value.ToString());
                txtComNumber.Text = DGVcom.CurrentRow.Cells[1].Value.ToString();                                                                                            //  dateTimePickerCom.Text = DGVcom.CurrentRow.Cells[2].Value.ToString();
                txtComModel.Text = DGVcom.CurrentRow.Cells[3].Value.ToString();
                txtComCustomer.Text = DGVcom.CurrentRow.Cells[4].Value.ToString();
                txtComAppearance.Text = DGVcom.CurrentRow.Cells[5].Value.ToString();
                CBComAppearanceSort.Text = DGVcom.CurrentRow.Cells[6].Value.ToString();
                txtComCause.Text = DGVcom.CurrentRow.Cells[7].Value.ToString();
                CBComCasueSort.Text = DGVcom.CurrentRow.Cells[8].Value.ToString();//原因分類
                CBComWarranty.Text = DGVcom.CurrentRow.Cells[9].Value.ToString();//保固
                txtComImNow.Text = DGVcom.CurrentRow.Cells[10].Value.ToString();//暫時對策
                CBComCur.Text = DGVcom.CurrentRow.Cells[11].Value.ToString();//目前進度
                txtImprovement.Text = DGVcom.CurrentRow.Cells[12].Value.ToString();//改善對策
                txtBulidPerson = DGVcom.CurrentRow.Cells[13].Value.ToString();//建單人
                txtComPerson.Text = DGVcom.CurrentRow.Cells[14].Value.ToString();//負責人
                                                                                 // dateTimePickerComFinish.Text = DGVcom.CurrentRow.Cells[15].Value.ToString();//預計完成日
                CBComDepartment.Text = DGVcom.CurrentRow.Cells[16].Value.ToString();
                txtStatus.Text = DGVcom.CurrentRow.Cells[17].Value.ToString();
                //18 建單日
                //19 實際完成日
                //2020/2/8  added
                if (DGVcom.CurrentRow.Cells[2].Value.ToString() == "" || DGVcom.CurrentRow.Cells[2].Value.ToString() == " ")
                {
                    dateTimePickerCom.Format = DateTimePickerFormat.Custom;
                    dateTimePickerCom.CustomFormat = " ";
                }
                else
                {
                    dateTimePickerCom.Format = DateTimePickerFormat.Long;
                    dateTimePickerCom.Text = DGVcom.CurrentRow.Cells[2].Value.ToString();
                }
                if (DGVcom.CurrentRow.Cells[15].Value.ToString() == "" || DGVcom.CurrentRow.Cells[15].Value.ToString() == " ")
                {
                    dateTimePickerComFinish.Format = DateTimePickerFormat.Custom;
                    dateTimePickerComFinish.CustomFormat = " ";
                }
                else
                {
                    dateTimePickerComFinish.Format = DateTimePickerFormat.Long;
                    dateTimePickerComFinish.Text = DGVcom.CurrentRow.Cells[15].Value.ToString();
                }

            }

            if (txtUserNameShow.Text == DGVcom.CurrentRow.Cells[13].Value.ToString())//非本人建立不得修改or delete
            {
                btnComSave.Text = "更新";
                btnComSave.Enabled = true;
                btnDelete.Enabled = true;
            }
            if (PlayRole == "FQC")//品管
            {
                btnComSave.Text = "更新";
                btnComSave.Enabled = true;
            }
        }

        private void DGVcom_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                DGVcom.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Beige;
                DGVcom.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Blue;
            }

        }

        private void btnComSave_Click(object sender, EventArgs e)
        {
            txtComAppearance.Text = txtComAppearance.Text.Replace("'", "\\'");//2022/11/10 can't 輸入報錯問題
                                                                              ////SQL  開始write //////////////////////////////////////////////////////////////////////////////////////////////    
            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();//開啟通道，建立連線，出現異常時,使用try catch語句
                using var cmd = new MySqlCommand();
                cmd.Connection = conn;
                if (bookComID == 0)//新增資料
                {
                    // 建單日 = DateTime.Now.ToLocalTime().ToString();//建單日時
                    cmd.CommandText = $"INSERT INTO rma.CustomerComplaint(客訴編號,客訴日期,機種型號," +
                      $"客戶名稱,保固,製品問題,問題現象分類,建單日,建單人)" +
                              "VALUES('" + txtComNumber.Text + "'" +
                              ",'" + dateTimePickerCom.Text + "'" +
                              ",'" + txtComModel.Text + "'" +
                              ",'" + txtComCustomer.Text + "'" +
                              ",'" + CBComWarranty.Text + "'" +
                              ",'" + txtComAppearance.Text + "'" +
                              ",'" + CBComAppearanceSort.Text + "'" +
                              ",'" + DateTime.Now.ToLocalTime().ToString() + "'" +
                              ",'" + txtBulidPerson + "')";
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("成功輸入資料!");
                }
                else//更新資料，只有建立者能更新or刪除
                {
                    switch (PlayRole)
                    {
                        case "sales":
                            cmd.CommandText = "UPDATE rma.CustomerComplaint SET 客訴編號='" + txtComNumber.Text + "'" +
                                    ",客訴日期='" + dateTimePickerCom.Text + "'" +
                                    ",機種型號='" + txtComModel.Text + "'" +
                                    ",客戶名稱='" + txtComCustomer.Text + "'" +
                                    ",製品問題='" + txtComAppearance.Text + "'" +
                                    ",保固='" + CBComWarranty.Text + "'" +
                                    ",問題現象分類='" + CBComAppearanceSort.Text + "'" +
                                    //  ",建單日='" + DateTime.Now.ToLocalTime().ToString() + "'" +
                                    //  ",建單人='" + txtBulidPerson + "'" +
                                    " WHERE 流水號='" + bookComID + "'";
                            cmd.ExecuteNonQuery();
                            conn.Close();
                            break;
                        case "FQC":
                            if (CBComCur.Text == "已完成")//完成日
                            {
                                cmd.CommandText = "UPDATE rma.CustomerComplaint SET 原因分析='" + txtComCause.Text + "'" +
                                ",問題原因分類='" + CBComCasueSort.Text + "'" +
                                ",暫時對策='" + txtComImNow.Text + "'" +
                                ",改善對策='" + txtImprovement.Text + "'" +
                                ",負責人='" + txtComPerson.Text + "'" +
                                ",責任單位='" + CBComDepartment.Text + "'" +
                                ",目前進度='" + CBComCur.Text + "'" +
                                ",預計完成日='" + dateTimePickerComFinish.Text + "'" +
                                ",實際完成日='" + DateTime.Now.ToLocalTime().ToString() + "'" +
                                " WHERE 流水號='" + bookComID + "'";
                            }
                            else
                            {
                                cmd.CommandText = "UPDATE rma.CustomerComplaint SET 原因分析='" + txtComCause.Text + "'" +
                                ",問題原因分類='" + CBComCasueSort.Text + "'" +
                                ",暫時對策='" + txtComImNow.Text + "'" +
                                ",改善對策='" + txtImprovement.Text + "'" +
                                ",負責人='" + txtComPerson.Text + "'" +
                                ",責任單位='" + CBComDepartment.Text + "'" +
                                ",目前進度='" + CBComCur.Text + "'" +
                                ",預計完成日='" + dateTimePickerComFinish.Text + "'" +
                                " WHERE 流水號='" + bookComID + "'";
                            }
                            cmd.ExecuteNonQuery();
                            conn.Close();
                            break;
                        default:
                            conn.Close();
                            break;
                    }
                    MessageBox.Show("成功更新資料!");
                }
                ClearCOM();
                GridFillCOM();//客訴頁面

            }
            catch (MySqlException ex)
            {
                MessageBox.Show("save SERVER連線異常!!!");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close();
            }

        }

        private void btnComUpload_Click(object sender, EventArgs e)
        {
            string directoryPath = @"\\192.168.1.3\public\品管品質\客訴單文件\" + txtComNumber.Text;
            // 檢查目錄是否存在，如果不存在就建立目錄
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
            // 設定檔案對話方塊的初始目錄
            openFileDialog1.InitialDirectory = directoryPath;

            // 允許選擇多個檔案
            openFileDialog1.Multiselect = true;
            openFileDialog1.FileName = "";

            // 打開檔案對話方塊並檢查使用者是否選擇了檔案
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // 獲取選擇的檔案清單
                string[] selectedFiles = openFileDialog1.FileNames;

                // 執行上傳檔案的相關邏輯
                foreach (string filePath in selectedFiles)
                {
                    // 可以在這裡進行檔案上傳的處理
                    // 您可以使用System.IO.Path類別來獲取檔案名稱等相關資訊

                    string fileName = Path.GetFileName(filePath);
                    string destinationPath = Path.Combine(directoryPath, fileName);

                    // 假設您使用一個自訂的檔案上傳函式UploadFile，將檔案從原始路徑上傳至目標路徑
                    UploadFile(filePath, destinationPath);

                    // 可以在這裡進行檔案上傳後的相關處理
                }
            }
        }
        private void UploadFile(string sourceFilePath, string destinationFilePath)
        {
            try
            {
                // 模擬檔案上傳的過程
                File.Copy(sourceFilePath, destinationFilePath);

                // 上傳成功後的相關處理
                MessageBox.Show("檔案上傳成功！");
            }
            catch (Exception ex)
            {
                // 上傳失敗的錯誤處理
                MessageBox.Show($"檔案上傳失敗：{ex.Message}");
            }
        }


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            System.Diagnostics.Process.Start("explorer.exe", @"\\192.168.1.3\public\佳文\軟體\RMA登錄工具");
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel3.LinkVisited = true;
            string filePath = @"\\192.168.1.3\public\品管品質\客訴單文件\" + txtComNumber.Text;

            if (System.IO.Directory.Exists(filePath))
            {
                System.Diagnostics.Process.Start("explorer.exe", filePath);
            }
            else
            {
                MessageBox.Show("Error: 尚未建立此客訴單的資料目錄", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnComOut_Click(object sender, EventArgs e)
        {
            ////////////////////////////////////////////////////////////////////////
            DataTable dt = new DataTable();
            //dt.Columns.Add("狀態");
            //dt.Columns.Add("維修校驗");
            //dt.Columns.Add("成品半成品");
            //dt.Columns.Add("機種名");
            //dt.Columns.Add("板名");
            foreach (DataGridViewColumn column in DGVcom.Columns)
                dt.Columns.Add(column.Name);
            for (int i = 0; i < DGVcom.Rows.Count; i++)
            {
                dt.Rows.Add();
                for (int j = 0; j < DGVcom.Columns.Count; j++)
                {
                    dt.Rows[i][j] = DGVcom.Rows[i].Cells[j].Value.ToString();
                }
            }
            MemoryStream stream = new MemoryStream();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            // 關閉新許可模式通知
            using (var ep1 = new ExcelPackage(stream))
            {
                var worksheet = ep1.Workbook.Worksheets.Add("客訴匯出");
                // ExcelWorksheet sheet = ep1.Workbook.Worksheets[0];
                //datatable 使用 LoadFromDatatable,collection 可使用 LoadFromCollection
                worksheet.Cells["A1"].LoadFromDataTable(dt, true);
                ep1.Save();
                string fileName = "客訴" + DateTime.Now.ToString("yyyyMMddHHmm");
                // string path1 = AppDomain.CurrentDomain.BaseDirectory + "xlsx\\"+fileName+".xlsx";
                // string filePath1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\xlsx\\setLotFile.xlsx";
                string path1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + fileName + ".xlsx";
                using FileStream myFile = File.Open(path1, FileMode.OpenOrCreate);
                stream.WriteTo(myFile);
                myFile.Close();
                MessageBox.Show("已成功存檔於桌面，檔名:" + fileName + ".xlsx");
            }
        } 
    }  
}







