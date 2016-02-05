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
using System.Data.OleDb;
using System.Data.Odbc;
using ClassLibrary;
using System.Windows;
using System.IO;
using MySql.Data.MySqlClient;
using MySql.Data;
using System.Collections;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Quartz;
using Quartz.Impl;
using System.Text.RegularExpressions;


namespace Volos
{
    public partial class Form1 : Form
    {
        private string connectionString;
        StringFormat strFormat; //Used to format the grid rows.
        ArrayList arrColumnLefts = new ArrayList();//Used to save left coordinates of columns
        ArrayList arrColumnWidths = new ArrayList();//Used to save column widths
        int iCellHeight = 0; //Used to get/set the datagridview cell height
        int iTotalWidth = 0; //
        int iRow = 0;//Used as counter
        bool bFirstPage = false; //Used to check whether we are printing first page
        bool bNewPage = false;// Used to check whether we are printing a new page
        int iHeaderHeight = 0; //Used for the header height

        public Form1()
        {
           InitializeComponent();

           dateTimePicker1.Format = DateTimePickerFormat.Custom;
           dateTimePicker1.CustomFormat = "yyyy-MM-dd HH:mm:ss.00";
           dateTimePicker1.Value = DateTime.Today;
           dateTimePicker2.Format = DateTimePickerFormat.Custom;
           dateTimePicker2.CustomFormat = "yyyy-MM-dd HH:mm:ss.00";
           dateTimePicker2.Value = DateTime.Today;

           string appPath = Path.GetDirectoryName(Application.ExecutablePath) + @"\tags.txt";
           //string appPath = @"C:\Documents and Settings\scada\Desktop\Wincc Statistics\tags.txt";

           List<ComboBoxPairs> cbp = new List<ComboBoxPairs>();
           List<string> lines = new List<string>();

           using (StreamReader r = new StreamReader(appPath, Encoding.Default))
           {
               string line;
               while ((line = r.ReadLine()) != null)
               {
                   lines.Add(line);
               }
           }

           foreach (string s in lines)
           {

               string[] words = Regex.Split(s,@"\\");

               cbp.Add(new ComboBoxPairs(words[1], words[0] + @"\" + words[1]));
               words[0] = "";
               words[1] = "";
           }

           comboBox1.DataSource = cbp;
           comboBox1.DisplayMember = "org";
           comboBox1.ValueMember = "org_latin";

           // construct a scheduler factory
           ISchedulerFactory schedFact = new StdSchedulerFactory();
           
           // get a scheduler
           IScheduler sched = schedFact.GetScheduler();
           sched.Start();

           // define the job and tie it to our HelloJob class
           IJobDetail job = JobBuilder.Create<HelloJob>()
               .WithIdentity("myJob", "group1")
               .Build();

           // Trigger the job to run now, and then every 40 seconds
           ITrigger trigger = TriggerBuilder.Create()
             .WithIdentity("myTrigger", "group1")
             .StartAt(DateBuilder.TomorrowAt(3, 0, 0))
             .WithSimpleSchedule(x => x
                 .WithIntervalInHours(24)
                 .RepeatForever())
             .Build();

           sched.ScheduleJob(job, trigger);

           //try
           //{
           //    MySqlConnection connection = new MySqlConnection(ClassLibrary.Class1.sqlstringtext());
           //    MySqlCommand command = connection.CreateCommand();
           //    MySqlDataReader reader;

           //    connection.Open();

           //    DateTime dt = DateTime.Today.AddDays(-1);

           //    command.CommandText = "SELECT * FROM calculations WHERE date='" + dt.ToString("yyyy-MM-dd") + "'";
           //    command.Prepare();
           //    reader = command.ExecuteReader();

           //    if (reader.Read() == false)
           //    {
           //        reader.Close();

           //        command.CommandText = "SELECT * FROM calculations  ORDER BY date DESC LIMIT 0,1";
           //        command.Prepare();
           //        reader = command.ExecuteReader();

           //        reader.Read();

           //        DateTime date = Convert.ToDateTime(reader["date"]);
           //        date = date.AddDays(1);
           //        while (date < DateTime.Today)
           //        {
           //            MessageBox.Show("Πράξεις σε εξέλιξη ....", "Φόρτωση", MessageBoxButtons.OK, MessageBoxIcon.Information);
           //            try
           //            {
           //                calculations(date);
           //            }
           //            catch (Exception exc)
           //            {
           //                MessageBox.Show(exc.Message, "Σφάλμα", MessageBoxButtons.OK, MessageBoxIcon.Error);
           //            }
           //            date = date.AddDays(1);
           //        }
           //    }

           //    reader.Close();

           //    connection.Close();
           //}
           //catch (Exception exc)
           //{
           //    MessageBox.Show(exc.Message, "Σφάλμα επικοινωνίας με διακομιστή", MessageBoxButtons.OK, MessageBoxIcon.Error);
           //}
        }

        public class ComboBoxPairs
        {
            public string org { get; set; }
            public string org_latin { get; set; }

            public ComboBoxPairs(string Org,
                                 string Org_latin)
            {
                org = Org;
                org_latin = Org_latin;
            }
        }

        public class HelloJob : IJob
        {
            public void Execute(IJobExecutionContext context)
            {

                try
                {
                    MySqlConnection connection = new MySqlConnection(ClassLibrary.Class1.sqlstringtext());
                    connection.Open();
                    MySqlCommand command = connection.CreateCommand();

                    string appPath = Path.GetDirectoryName(Application.ExecutablePath) + @"\tags.txt";
                    //string appPath = @"c:\Users\" + Environment.UserName + @"\Desktop\sql.txt";
                    List<string> lines = new List<string>();
                    DateTime date = DateTime.Today.AddDays(-1);

                    using (StreamReader r = new StreamReader(appPath, Encoding.Default))
                    {
                        string line;
                        while ((line = r.ReadLine()) != null)
                        {
                            lines.Add(line);
                        }
                    }
                    List<string> tags = new List<string>();
                    foreach (string s in lines)
                    {
                        tags.Add(Convert.ToString(ClassLibrary.Class1.integral(s.Trim(), date)));
                    }

                    command.CommandText = "insert into calculations(A6_FLOW,A3_FLOW,A5_FLOW,PAROXH_A_ILYOS_1,PAROXH_A_ILYOS_3,PAROXH_PROP_1,PAROXH_PROP_2,PERISIA_PAROXI,25IF02,FLOW_OMOG,FLOW,MLSS_A_ILYOS_1,MLSS_A_ILYOS_3,MLSS_1,MLSS_2,MLSS_3,MLSS_4,MLSS_5,PAROXI_EISODOU,SUM,MLSS,MLSS_PERIS_NEW,IDO_1,IDO_2,IDO_3,IDO_4,IDO_05,MLSS_DEC,BLOWERS_POWER,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13, @p14, @p15, @p16, @p17, @p18, @p19, @p20, @p21, @p22, @p23, @p24, @p25, @p26, @p27, @p28, @p29, @p30)";
                    command.Prepare();
                    command.Parameters.AddWithValue("@p1", tags[0]);
                    command.Parameters.AddWithValue("@p2", tags[1]);
                    command.Parameters.AddWithValue("@p3", tags[2]);
                    command.Parameters.AddWithValue("@p4", tags[3]);
                    command.Parameters.AddWithValue("@p5", tags[4]);
                    command.Parameters.AddWithValue("@p6", tags[5]);
                    command.Parameters.AddWithValue("@p7", tags[6]);
                    command.Parameters.AddWithValue("@p8", tags[7]);
                    command.Parameters.AddWithValue("@p9", tags[8]);
                    command.Parameters.AddWithValue("@p10", tags[9]);
                    command.Parameters.AddWithValue("@p11", tags[10]);
                    command.Parameters.AddWithValue("@p12", tags[11]);
                    command.Parameters.AddWithValue("@p13", tags[12]);
                    command.Parameters.AddWithValue("@p14", tags[13]);
                    command.Parameters.AddWithValue("@p15", tags[14]);
                    command.Parameters.AddWithValue("@p16", tags[15]);
                    command.Parameters.AddWithValue("@p17", tags[16]);
                    command.Parameters.AddWithValue("@p18", tags[17]);
                    command.Parameters.AddWithValue("@p19", tags[18]);
                    command.Parameters.AddWithValue("@p20", tags[19]);
                    command.Parameters.AddWithValue("@p21", tags[20]);
                    command.Parameters.AddWithValue("@p22", tags[21]);
                    command.Parameters.AddWithValue("@p23", tags[22]);
                    command.Parameters.AddWithValue("@p24", tags[23]);
                    command.Parameters.AddWithValue("@p25", tags[24]);
                    command.Parameters.AddWithValue("@p26", tags[25]);
                    command.Parameters.AddWithValue("@p27", tags[26]);
                    command.Parameters.AddWithValue("@p28", tags[27]);
                    command.Parameters.AddWithValue("@p29", tags[28]);
                    command.Parameters.AddWithValue("@p30", date);
                    command.ExecuteNonQuery();

                    ClassLibrary.Class1.eisodos_eel(date);
                    ClassLibrary.Class1.a_kath(date);
                    ClassLibrary.Class1.dex_aer(date);
                    ClassLibrary.Class1.b_kath(date);
                    ClassLibrary.Class1.c_vathm(date);
                    ClassLibrary.Class1.ol_apod_eel(date);
                    ClassLibrary.Class1.propax(date);
                    ClassLibrary.Class1.mhx_pax(date);
                    ClassLibrary.Class1.xwneusi(date);
                    ClassLibrary.Class1.metapax(date);
                    ClassLibrary.Class1.afudat(date);

                    MessageBox.Show("Οι πράξεις ολοκληρώθηκαν");

                    connection.Close();
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message, "Σφάλμα επικοινωνίας με διακομιστή", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void change_interval(int seconds,int minutes,int hour,int dayofmonth,int month,int year)
        {
            IScheduler sched = new StdSchedulerFactory().GetScheduler();
            
            // Define a new Trigger 
            ITrigger trigger = TriggerBuilder.Create()
             .WithIdentity("myTrigger", "group1")
             .StartAt(DateBuilder.DateOf(hour,minutes,seconds,dayofmonth,month,year))
             .WithSimpleSchedule(x => x
                 .WithIntervalInHours(24)
                 .RepeatForever())
             .Build();

            TriggerKey triggerkey = new TriggerKey("myTrigger", "group1");

            // tell the scheduler to remove the old trigger with the given key, and put the new one in its place
            sched.RescheduleJob(triggerkey, trigger);
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (dateTimePicker2.Visible == false)
            {
                try
                {
                    MySqlConnection connection1 = new MySqlConnection(ClassLibrary.Class1.sqlstringtext());
                    connection1.Open();
                    MySqlCommand command = connection1.CreateCommand();
                    MySqlDataReader reader;

                    command.CommandText = "SELECT * FROM eisodos_eel WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox1.Text = reader["straggidia"].ToString();
                    textBox2.Text = reader["propax"].ToString();
                    textBox3.Text = reader["mhx_pax"].ToString();
                    textBox4.Text = reader["metapax"].ToString();
                    textBox5.Text = reader["afudat"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM a_kath WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox6.Text = reader["xronos_param"].ToString();
                    textBox7.Text = reader["paroxi_a_il"].ToString();
                    textBox8.Text = reader["ssa"].ToString();
                    textBox9.Text = reader["fortio_ssa"].ToString();
                    textBox10.Text = reader["fortio_ptot_in"].ToString();
                    textBox11.Text = reader["fortio_ptot_out"].ToString();
                    textBox12.Text = reader["krok_fe"].ToString();
                    textBox13.Text = reader["katan_fe"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM b_kath WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox14.Text = reader["sun_par"].ToString();
                    textBox15.Text = reader["udrau_fort"].ToString();
                    textBox16.Text = reader["fort_ster"].ToString();
                    textBox17.Text = reader["xron_param"].ToString();
                    textBox18.Text = reader["tax_ex"].ToString();
                    textBox19.Text = reader["fort_ss_out"].ToString();
                    textBox20.Text = reader["bod_apom"].ToString();
                    textBox21.Text = reader["bod_apod"].ToString();
                    textBox22.Text = reader["ss_apod"].ToString();
                    textBox23.Text = reader["fort_n_nh3_out"].ToString();
                    textBox24.Text = reader["fort_n_no3_out"].ToString();
                    textBox25.Text = reader["n_nitro"].ToString();
                    textBox26.Text = reader["n_aponitro"].ToString();
                    textBox27.Text = reader["fort_ptot_out"].ToString();
                    textBox28.Text = reader["katan_fe"].ToString();
                    textBox29.Text = reader["apod_ptot_b"].ToString();
                    textBox30.Text = reader["apod_ptot_a"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM ol_apod_eel WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox31.Text = reader["apod_ptot"].ToString();
                    textBox32.Text = reader["apod_n_nh3"].ToString();
                    textBox33.Text = reader["apod_n_no3"].ToString();
                    textBox34.Text = reader["apod_bod"].ToString();
                    textBox35.Text = reader["apod_ss"].ToString();
                    textBox36.Text = reader["apod_ntot"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM dex_aer WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox37.Text = reader["xronos_param_aer"].ToString();
                    textBox38.Text = reader["pos_ptht"].ToString();
                    textBox39.Text = reader["ol_mlss_aer"].ToString();
                    textBox40.Text = reader["ol_mlvss_aer"].ToString();
                    textBox41.Text = reader["par_aera"].ToString();
                    textBox42.Text = reader["logos_aera"].ToString();
                    textBox43.Text = reader["bod_apom"].ToString();
                    textBox44.Text = reader["logosaer_pr_bodapom"].ToString();
                    textBox45.Text = reader["apod_o2"].ToString();
                    textBox46.Text = reader["par_o2"].ToString();
                    textBox47.Text = reader["logoso2_pr_bodapom"].ToString();
                    textBox48.Text = reader["ogk_fort_bod"].ToString();
                    textBox49.Text = reader["sun_pax_aer"].ToString();
                    textBox50.Text = reader["vss_upol"].ToString();
                    textBox51.Text = reader["inerts_upol"].ToString();
                    textBox52.Text = reader["tss_upol"].ToString();
                    textBox53.Text = reader["tss_diad"].ToString();
                    textBox54.Text = reader["mcrt"].ToString();
                    textBox55.Text = reader["f_m"].ToString();
                    textBox56.Text = reader["n_nh3_eis"].ToString();
                    textBox57.Text = reader["n_no3_eis"].ToString();
                    textBox58.Text = reader["n_nh3_ex"].ToString();
                    textBox59.Text = reader["n_no3_ex"].ToString();
                    textBox60.Text = reader["n_nitro"].ToString();
                    textBox61.Text = reader["n_aponitro"].ToString();
                    textBox62.Text = reader["xron_aer"].ToString();
                    textBox63.Text = reader["xron_aponitro"].ToString();
                    textBox64.Text = reader["apod_n_nh3"].ToString();
                    textBox65.Text = reader["apod_n_no3"].ToString();
                    textBox66.Text = reader["fort_ptot_out"].ToString();
                    textBox67.Text = reader["xron_param"].ToString();
                    textBox68.Text = reader["apod_ptot"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM c_vathm WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox69.Text = reader["udrau_fort"].ToString();
                    textBox70.Text = reader["fort_ster"].ToString();
                    textBox71.Text = reader["apod_filtr"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM propax WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox72.Text = reader["xron_param"].ToString();
                    textBox73.Text = reader["udrau_fort"].ToString();
                    textBox74.Text = reader["fort_ster"].ToString();
                    textBox75.Text = reader["fort_ss_pax"].ToString();
                    textBox76.Text = reader["fort_ss_stag"].ToString();
                    textBox77.Text = reader["ss_stag"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM mhx_pax WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox78.Text = reader["fort_wass"].ToString();
                    textBox79.Text = reader["par_omog"].ToString();
                    textBox80.Text = reader["fort_ss_pax"].ToString();
                    textBox81.Text = reader["fort_ss_strag"].ToString();
                    textBox82.Text = reader["ss_strag"].ToString();
                    textBox83.Text = reader["eid_kat_phl"].ToString();
                    textBox84.Text = reader["logos_pax"].ToString();
                    textBox85.Text = reader["fort_omog"].ToString();
                    textBox86.Text = reader["ss_pax"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM xwneusi WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox87.Text = reader["xron_param"].ToString();
                    textBox88.Text = reader["fort_ss_xwn"].ToString();
                    textBox89.Text = reader["fort_vss_xwn"].ToString();
                    textBox90.Text = reader["pos_ptht_in"].ToString();
                    textBox91.Text = reader["fort_ptht"].ToString();
                    textBox92.Text = reader["fort_ss_fix"].ToString();
                    textBox93.Text = reader["fort_ss_ex"].ToString();
                    textBox94.Text = reader["fort_vss_ex"].ToString();
                    textBox95.Text = reader["pos_ptht_out"].ToString();
                    textBox96.Text = reader["parag_vioaer"].ToString();
                    textBox97.Text = reader["apodosi"].ToString();
                    textBox98.Text = reader["logos_va"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM metapax WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox99.Text = reader["xron_param"].ToString();
                    textBox100.Text = reader["udrau_fort"].ToString();
                    textBox101.Text = reader["fort_ster"].ToString();
                    textBox102.Text = reader["fort_ss_pax"].ToString();
                    textBox103.Text = reader["fort_ss_strag"].ToString();
                    textBox104.Text = reader["ss_strag"].ToString();

                    reader.Close();

                    command.CommandText = "SELECT * FROM afudatwsi WHERE date='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                    command.Prepare();
                    reader = command.ExecuteReader();

                    reader.Read();

                    textBox105.Text = reader["fort_ss_afud"].ToString();
                    textBox106.Text = reader["fort_ss_strag"].ToString();
                    textBox107.Text = reader["eid_kat_phl"].ToString();

                    reader.Close();

                    connection1.Close();
                }
                catch (Exception)
                {
                    MessageBox.Show("Ανάκτηση Δεδομένων .... Περιμένετε ....", "Φόρτωση", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    try
                    {
                        calculations(dateTimePicker1.Value.Date);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Δεν υπάρχουν καταγεγραμμένες τιμές για την συγκεκριμένη ημερομηνία", "Ειδοποίηση", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
            else
            {
                if (comboBox1.Visible == true)
                {
                    if (comboBox1.SelectedItem != null)
                    {
                        try
                        {
                            connectionString = ClassLibrary.Class1.winccstringtext();

                            DataSet ds = new DataSet();

                            OleDbConnection connection;
                            connection = new OleDbConnection(connectionString);

                            OleDbDataAdapter adapter;

                            DateTime dtime1 = dateTimePicker1.Value.AddHours(-3);
                            DateTime dtime2 = dateTimePicker2.Value.AddHours(-3);

                            string sql = null;
                            sql = @"Tag:R,'" + comboBox1.SelectedValue + "','" + dtime1.ToString("yyyy-MM-dd HH:mm:ss.00") + "','" + dtime2.ToString("yyyy-MM-dd HH:mm:ss.00") + "'";

                            connection.Open();

                            adapter = new OleDbDataAdapter(sql, connection);
                            adapter.Fill(ds);

                            DataTable table = ds.Tables[0];

                            DataRow tempRow = null;
                            int j = 0;
                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;

                                DateTime date = new DateTime();
                                date = Convert.ToDateTime(tempRow["Timestamp"]);
                                date = date.AddHours(3);
                                table.Rows[j][1] = date.ToString();
                                j++;
                            }

                            ds.Tables.Clear();
                            ds.Tables.Add(table);

                            dataGridView1.DataSource = ds.Tables[0];

                            connection.Close();

                            int sum = 0;
                            for (int i = 0; i < table.Rows.Count; ++i)
                            {
                                sum += Convert.ToInt32(table.Rows[i][2]);
                            }

                            DateTime dt = new DateTime();
                            DateTime dt1 = new DateTime();
                            dt = Convert.ToDateTime(table.Rows[0][1]);
                            dt1 = Convert.ToDateTime(table.Rows[1][1]);
                            double step = (dt1 - dt).TotalHours;

                            double result = sum * step;

                            label4.Text = "ΜΕΣΟΣ ΟΡΟΣ TAG";
                            label5.Text = Convert.ToString(result);
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message, "Σφάλμα επικοινωνίας με διακομιστή", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Επιλέξτε ένα TAG πρώτα");
                    }
                }
                else
                {
                    try
                    {
                        if (checkBox1.Checked == false)
                        {
                            MySqlConnection connection1 = new MySqlConnection(ClassLibrary.Class1.sqlstringtext());

                            DataSet ds1 = new DataSet();

                            string com = "SELECT * FROM eisodos_eel WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            MySqlDataAdapter adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "eisodos_eel");

                            DataTable table = ds1.Tables["eisodos_eel"];

                            decimal[] sum = new decimal[40];
                            DataRow tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                            }
                            DataRow row = (DataRow)table.NewRow();

                            row[0] = sum[0];
                            row[1] = sum[1];
                            row[2] = sum[2];
                            row[3] = sum[3];
                            row[4] = sum[4];
                            
                            table.Rows.Add(row);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView2.DataSource = table;
                            dataGridView2.Columns["date"].DisplayIndex = 0;
                            dataGridView2.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView2.Columns[0].HeaderText = "Στραγγίδια (m³/d)";
                            dataGridView2.Columns[1].HeaderText = "Προπάχυνση (m³/d)";
                            dataGridView2.Columns[2].HeaderText = "Μηχανική Πάχυνση (m³/d)";
                            dataGridView2.Columns[3].HeaderText = "Μεταπάχυνση (m³/d)";
                            dataGridView2.Columns[4].HeaderText = "Αφυδάτωση (m³/d)";

                            dataGridView2.Rows[dataGridView2.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView2.Rows[dataGridView2.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                            
                            com = "SELECT * FROM a_kath WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "a_kath");

                            DataTable table1 = ds1.Tables["a_kath"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table1.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                                sum[5] += Convert.ToDecimal(tempRow[5]);
                                sum[6] += Convert.ToDecimal(tempRow[6]);
                                sum[7] += Convert.ToDecimal(tempRow[7]);
                            }

                            DataRow row1 = (DataRow)table1.NewRow();

                            row1[0] = sum[0];
                            row1[1] = sum[1];
                            row1[2] = sum[2];
                            row1[3] = sum[3];
                            row1[4] = sum[4];
                            row1[5] = sum[5];
                            row1[6] = sum[6];
                            row1[7] = sum[7];
                            table1.Rows.Add(row1);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView3.DataSource = table1;
                            dataGridView3.Columns["date"].DisplayIndex = 0;
                            dataGridView3.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView3.Columns[0].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView3.Columns[1].HeaderText = "Παροχή Α' Ιλύς (m³/d)";
                            dataGridView3.Columns[2].HeaderText = "SS a (mg/l)";
                            dataGridView3.Columns[3].HeaderText = "Φορτίο SS a (Kg/d)";
                            dataGridView3.Columns[4].HeaderText = "Φορτίο Ptot in (Kg/d)";
                            dataGridView3.Columns[5].HeaderText = "Φορτίο Ptot out (Kg/d)";
                            dataGridView3.Columns[6].HeaderText = "Κροκιδωτικό FE (Kg/d)";
                            dataGridView3.Columns[7].HeaderText = "Κατανάλωση FE/Ptot Απομ. (Kg/Kg)";

                            dataGridView3.Rows[dataGridView3.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView3.Rows[dataGridView3.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM b_kath WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "b_kath");

                            table = ds1.Tables["b_kath"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                                sum[5] += Convert.ToDecimal(tempRow[5]);
                                sum[6] += Convert.ToDecimal(tempRow[6]);
                                sum[7] += Convert.ToDecimal(tempRow[7]);
                                sum[8] += Convert.ToDecimal(tempRow[8]);
                                sum[9] += Convert.ToDecimal(tempRow[9]);
                                sum[10] += Convert.ToDecimal(tempRow[10]);
                                sum[11] += Convert.ToDecimal(tempRow[11]);
                                sum[12] += Convert.ToDecimal(tempRow[12]);
                                sum[13] += Convert.ToDecimal(tempRow[13]);
                                sum[14] += Convert.ToDecimal(tempRow[14]);
                                sum[15] += Convert.ToDecimal(tempRow[15]);
                                sum[16] += Convert.ToDecimal(tempRow[16]);
                            }

                            DataRow row2 = (DataRow)table.NewRow();

                            row2[0] = sum[0];
                            row2[1] = sum[1];
                            row2[2] = sum[2];
                            row2[3] = sum[3];
                            row2[4] = sum[4];
                            row2[5] = sum[5];
                            row2[6] = sum[6];
                            row2[7] = sum[7];
                            row2[8] = sum[8];
                            row2[9] = sum[9];
                            row2[10] = sum[10];
                            row2[11] = sum[11];
                            row2[12] = sum[12];
                            row2[13] = sum[13];
                            row2[14] = sum[14];
                            row2[15] = sum[15];
                            row2[16] = sum[16];
                            table.Rows.Add(row2);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView4.DataSource = table;
                            dataGridView4.Columns["date"].DisplayIndex = 0;
                            dataGridView4.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView4.Columns[0].HeaderText = "Συνολική Παροχή (m³/d)";
                            dataGridView4.Columns[1].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView4.Columns[2].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView4.Columns[3].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView4.Columns[4].HeaderText = "Ταχύτητα Εξ. (m²/d)";
                            dataGridView4.Columns[5].HeaderText = "Φορτίο SS out (Kg/d)";
                            dataGridView4.Columns[6].HeaderText = "BOD Απομακρυνόμενο (Kg/d)";
                            dataGridView4.Columns[7].HeaderText = "BOD Απόδοση (%)";
                            dataGridView4.Columns[8].HeaderText = "SS Απόδοση (%)";
                            dataGridView4.Columns[9].HeaderText = "Φορτιο Ν-ΝΗ3 out (Kg/d)";
                            dataGridView4.Columns[10].HeaderText = "Φορτιο Ν-ΝΟ3 out (Kg/d)";
                            dataGridView4.Columns[11].HeaderText = "N Νιτροποίησης (Kg/d)";
                            dataGridView4.Columns[12].HeaderText = "N Απονιτροποίησης (Kg/d)";
                            dataGridView4.Columns[13].HeaderText = "Φορτίο Ptot out (Kg/d)";
                            dataGridView4.Columns[14].HeaderText = "Κατανάλωση Fe/απομακρυνόμενου P (Kg/Kg)";
                            dataGridView4.Columns[15].HeaderText = "Απόδοση απομ. Ptot β' καθ. (%)";
                            dataGridView4.Columns[16].HeaderText = "Απόδοση απομ. Ptot α' καθ. (%)";

                            dataGridView4.Rows[dataGridView4.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView4.Rows[dataGridView4.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM ol_apod_eel WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "ol_apod_eel");

                            table = ds1.Tables["ol_apod_eel"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                                sum[5] += Convert.ToDecimal(tempRow[5]);
                            }

                            DataRow row3 = (DataRow)table.NewRow();

                            row3[0] = sum[0];
                            row3[1] = sum[1];
                            row3[2] = sum[2];
                            row3[3] = sum[3];
                            row3[4] = sum[4];
                            row3[5] = sum[5];

                            table.Rows.Add(row3);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView5.DataSource = table;
                            dataGridView5.Columns["date"].DisplayIndex = 0;
                            dataGridView5.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView5.Columns[0].HeaderText = "Απόδοση απομ. Ptot (%)";
                            dataGridView5.Columns[1].HeaderText = "Απόδοση απομ. N-NH3 (%)";
                            dataGridView5.Columns[2].HeaderText = "Απόδοση απομ. N-NO3 (%)";
                            dataGridView5.Columns[3].HeaderText = "Απόδοση απομ. BOD (%)";
                            dataGridView5.Columns[4].HeaderText = "Απόδοση απομ. SS (%)";
                            dataGridView5.Columns[5].HeaderText = "Απόδοση απομ. Ntot (%)";

                            dataGridView5.Rows[dataGridView5.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView5.Rows[dataGridView5.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM dex_aer WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "dex_aer");

                            table = ds1.Tables["dex_aer"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 32; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row4 = (DataRow)table.NewRow();

                            for (int i = 0; i < 32; i++)
                            {
                                row4[i] = sum[i];
                            }

                            table.Rows.Add(row4);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView6.DataSource = table;
                            dataGridView6.Columns["date"].DisplayIndex = 0;
                            dataGridView6.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView6.Columns[0].HeaderText = "Χρόνος Παραμονής στον αερ. (h)";
                            dataGridView6.Columns[1].HeaderText = "Ποσοστό Πτητικών (%)";
                            dataGridView6.Columns[2].HeaderText = "Ολικά MLSS αερ. (Kg)";
                            dataGridView6.Columns[3].HeaderText = "Ολικά MLVSS αερ. (Kg)";
                            dataGridView6.Columns[4].HeaderText = "Παροχή αέρα (m³/d)";
                            dataGridView6.Columns[5].HeaderText = "Λόγος αέρα προς παρ. εισ. (m³/m³)";
                            dataGridView6.Columns[6].HeaderText = "BOD απομακρυνόμενο (m³/Kg)";
                            dataGridView6.Columns[7].HeaderText = "Λόγος αέρα προς απομακρυνόμενο BOD (Kg/m³)";
                            dataGridView6.Columns[8].HeaderText = "Απόδοση σε Ο2 (Kg/m³)";
                            dataGridView6.Columns[9].HeaderText = "Παροχή O2 (Kg/d)";
                            dataGridView6.Columns[10].HeaderText = "Λόγος O2 προς απομακρυνόμενο BOD (Kg/Kg)";
                            dataGridView6.Columns[11].HeaderText = "Oγκομετρικό φορτίο BOD (Kg/d*m³)";
                            dataGridView6.Columns[12].HeaderText = "Συνολική Παροχή Αέρα (m³/d)";
                            dataGridView6.Columns[13].HeaderText = "VSS υπολογισμός (Kg/d)";
                            dataGridView6.Columns[14].HeaderText = "Inerts υπολογισμός (Kg/d)";
                            dataGridView6.Columns[15].HeaderText = "TSS υπολογισμός (Kg/d)";
                            dataGridView6.Columns[16].HeaderText = "TSS στη διαδικασία (Kg)";
                            dataGridView6.Columns[17].HeaderText = "MCRT (d)";
                            dataGridView6.Columns[18].HeaderText = "F/M (Kg*d/Kg)";
                            dataGridView6.Columns[19].HeaderText = "Φορτίο Ν-ΝΗ3 εισ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[20].HeaderText = "Φορτίο Ν-ΝΟ3 εισ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[21].HeaderText = "Φορτίο Ν-ΝΗ3 εξ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[22].HeaderText = "Φορτίο Ν-ΝΟ3 εξ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[23].HeaderText = "N Νιτροποίησης (Kg/d)";
                            dataGridView6.Columns[24].HeaderText = "N Απονιτροποίησης (Kg/d)";
                            dataGridView6.Columns[25].HeaderText = "Χρόνος Αερισμού (h)";
                            dataGridView6.Columns[26].HeaderText = "Χρονος Απονιτροποίησης (h)";
                            dataGridView6.Columns[27].HeaderText = "Απόδοση απομ. Ν-ΝΗ3 (%)";
                            dataGridView6.Columns[28].HeaderText = "Απόδοση απομ. Ν-ΝO3 (%)";
                            dataGridView6.Columns[29].HeaderText = "Φορτίο Ptot out (Kg/d)";
                            dataGridView6.Columns[30].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView6.Columns[31].HeaderText = "Απόδοση απομ. Ptot (%)";

                            dataGridView6.Rows[dataGridView6.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView6.Rows[dataGridView6.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM c_vathm WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "c_vathm");

                            table = ds1.Tables["c_vathm"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 3; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row5 = (DataRow)table.NewRow();

                            for (int i = 0; i < 3; i++)
                            {
                                row5[i] = sum[i];
                            }

                            table.Rows.Add(row5);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView7.DataSource = ds1.Tables["c_vathm"];
                            dataGridView7.Columns["date"].DisplayIndex = 0;
                            dataGridView7.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView7.Columns[0].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView7.Columns[1].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView7.Columns[2].HeaderText = "Απόδοση Φίλτρων (%)";

                            dataGridView7.Rows[dataGridView7.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView7.Rows[dataGridView7.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM propax WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "propax");

                            table = ds1.Tables["propax"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 6; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row6 = (DataRow)table.NewRow();

                            for (int i = 0; i < 6; i++)
                            {
                                row6[i] = sum[i];
                            }

                            table.Rows.Add(row6);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView8.DataSource = ds1.Tables["propax"];
                            dataGridView8.Columns["date"].DisplayIndex = 0;
                            dataGridView8.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView8.Columns[0].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView8.Columns[1].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView8.Columns[2].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView8.Columns[3].HeaderText = "Φορτίο SS παχυμένης (kg/d)";
                            dataGridView8.Columns[4].HeaderText = "Φορτίο SS στραγγιδίων (kg/d)";
                            dataGridView8.Columns[5].HeaderText = "SS στραγγιδίων (mg/l)";

                            dataGridView8.Rows[dataGridView8.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView8.Rows[dataGridView8.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM mhx_pax WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "mhx_pax");

                            table = ds1.Tables["mhx_pax"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 9; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row7 = (DataRow)table.NewRow();

                            for (int i = 0; i < 9; i++)
                            {
                                row7[i] = sum[i];
                            }

                            table.Rows.Add(row7);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView9.DataSource = ds1.Tables["mhx_pax"];
                            dataGridView9.Columns["date"].DisplayIndex = 0;
                            dataGridView9.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView9.Columns[0].HeaderText = "Φορτίο WASS (kg/d)";
                            dataGridView9.Columns[1].HeaderText = "Παροχή προς ομογ. (m³/d)";
                            dataGridView9.Columns[2].HeaderText = "Φορτίο SS παχυμ. (kg/d)";
                            dataGridView9.Columns[3].HeaderText = "Φορτίο SS στραγγ. (kg/d)";
                            dataGridView9.Columns[4].HeaderText = "SS στραγγ. (mg/l)";
                            dataGridView9.Columns[5].HeaderText = "Ειδική καταν. ΠΗΛ (kg/Kg)";
                            dataGridView9.Columns[6].HeaderText = "Λόγος Πάχυνσης (decimal)";
                            dataGridView9.Columns[7].HeaderText = "Φορτίο ομογεν. (kg/d)";
                            dataGridView9.Columns[8].HeaderText = "SS παχυμ. (mg/l)";

                            dataGridView9.Rows[dataGridView9.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView9.Rows[dataGridView9.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM xwneusi WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "xwneusi");

                            table = ds1.Tables["xwneusi"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 12; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row8 = (DataRow)table.NewRow();

                            for (int i = 0; i < 12; i++)
                            {
                                row8[i] = sum[i];
                            }

                            table.Rows.Add(row8);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView10.DataSource = ds1.Tables["xwneusi"];
                            dataGridView10.Columns["date"].DisplayIndex = 0;
                            dataGridView10.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView10.Columns[0].HeaderText = "Χρόνος Παραμονής (d)";
                            dataGridView10.Columns[1].HeaderText = "Φορτίο SS χώνευσης (kg/d)";
                            dataGridView10.Columns[2].HeaderText = "Φορτίο VSS χώνευσης (kg/d)";
                            dataGridView10.Columns[3].HeaderText = "Ποσοστό Πτητικών in (%)";
                            dataGridView10.Columns[4].HeaderText = "Φόρτιση Πτητικών (kg/d*m³)";
                            dataGridView10.Columns[5].HeaderText = "Φορτίο SS fix (kg/d)";
                            dataGridView10.Columns[6].HeaderText = "Φορτίο SS εξόδου (kg/d)";
                            dataGridView10.Columns[7].HeaderText = "Φορτίο VSS εξόδου (kg/d)";
                            dataGridView10.Columns[8].HeaderText = "Ποσοστό Πτητικών out (%)";
                            dataGridView10.Columns[9].HeaderText = "Παραγωγή Βιοαερίου (kg/d)";
                            dataGridView10.Columns[10].HeaderText = "Απόδοση (%)";
                            dataGridView10.Columns[11].HeaderText = "Λόγος VA/A (decimal)";

                            dataGridView10.Rows[dataGridView10.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView10.Rows[dataGridView10.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM metapax WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "metapax");

                            table = ds1.Tables["metapax"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 6; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row9 = (DataRow)table.NewRow();

                            for (int i = 0; i < 6; i++)
                            {
                                row9[i] = sum[i];
                            }

                            table.Rows.Add(row9);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView11.DataSource = ds1.Tables["metapax"];
                            dataGridView11.Columns["date"].DisplayIndex = 0;
                            dataGridView11.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView11.Columns[0].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView11.Columns[1].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView11.Columns[2].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView11.Columns[3].HeaderText = "Φορτίο SS παχυμένης (kg/d)";
                            dataGridView11.Columns[4].HeaderText = "Φορτίο SS στραγγιδίων (kg/d)";
                            dataGridView11.Columns[5].HeaderText = "SS στραγγιδίων (mg/l)";

                            dataGridView11.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView11.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM afudatwsi WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "afudatwsi");

                            table = ds1.Tables["afudatwsi"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 3; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row10 = (DataRow)table.NewRow();

                            for (int i = 0; i < 3; i++)
                            {
                                row10[i] = sum[i];
                            }

                            table.Rows.Add(row10);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView12.DataSource = ds1.Tables["afudatwsi"];
                            dataGridView12.Columns["date"].DisplayIndex = 0;
                            dataGridView12.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView12.Columns[0].HeaderText = "Φορτίο SS αφυδατωμένης (kg/d)";
                            dataGridView12.Columns[1].HeaderText = "Φορτίο SS στραγγιδίων (kg/d)";
                            dataGridView12.Columns[2].HeaderText = "Eιδική καταν. ΠΗΛ (kg/Kg)";

                            dataGridView12.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView12.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                        }
                        else
                        {
                            MySqlConnection connection1 = new MySqlConnection(ClassLibrary.Class1.sqlstringtext());

                            DataSet ds1 = new DataSet();

                            string com = "SELECT * FROM eisodos_eel WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            MySqlDataAdapter adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "eisodos_eel");

                            DataTable table = ds1.Tables["eisodos_eel"];

                            decimal[] sum = new decimal[40];
                            DataRow tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                            }
                            DataRow row = (DataRow)table.NewRow();

                            row[0] = sum[0];
                            row[1] = sum[1];
                            row[2] = sum[2];
                            row[3] = sum[3];
                            row[4] = sum[4];
                            table.Rows.Clear();
                            table.Rows.Add(row);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView2.DataSource = table;
                            dataGridView2.Columns["date"].DisplayIndex = 0;
                            dataGridView2.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView2.Columns[0].HeaderText = "Στραγγίδια (m³/d)";
                            dataGridView2.Columns[1].HeaderText = "Προπάχυνση (m³/d)";
                            dataGridView2.Columns[2].HeaderText = "Μηχανική Πάχυνση (m³/d)";
                            dataGridView2.Columns[3].HeaderText = "Μεταπάχυνση (m³/d)";
                            dataGridView2.Columns[4].HeaderText = "Αφυδάτωση (m³/d)";

                            dataGridView2.Rows[dataGridView2.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView2.Rows[dataGridView2.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM a_kath WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "a_kath");

                            DataTable table1 = ds1.Tables["a_kath"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table1.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                                sum[5] += Convert.ToDecimal(tempRow[5]);
                                sum[6] += Convert.ToDecimal(tempRow[6]);
                                sum[7] += Convert.ToDecimal(tempRow[7]);
                            }

                            DataRow row1 = (DataRow)table1.NewRow();

                            row1[0] = sum[0];
                            row1[1] = sum[1];
                            row1[2] = sum[2];
                            row1[3] = sum[3];
                            row1[4] = sum[4];
                            row1[5] = sum[5];
                            row1[6] = sum[6];
                            row1[7] = sum[7];
                            table1.Rows.Clear();
                            table1.Rows.Add(row1);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView3.DataSource = table1;
                            dataGridView3.Columns["date"].DisplayIndex = 0;
                            dataGridView3.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView3.Columns[0].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView3.Columns[1].HeaderText = "Παροχή Α' Ιλύς (m³/d)";
                            dataGridView3.Columns[2].HeaderText = "SS a (mg/l)";
                            dataGridView3.Columns[3].HeaderText = "Φορτίο SS a (Kg/d)";
                            dataGridView3.Columns[4].HeaderText = "Φορτίο Ptot in (Kg/d)";
                            dataGridView3.Columns[5].HeaderText = "Φορτίο Ptot out (Kg/d)";
                            dataGridView3.Columns[6].HeaderText = "Κροκιδωτικό FE (Kg/d)";
                            dataGridView3.Columns[7].HeaderText = "Κατανάλωση FE/Ptot Απομ. (Kg/Kg)";

                            dataGridView3.Rows[dataGridView3.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView3.Rows[dataGridView3.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM b_kath WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "b_kath");

                            table = ds1.Tables["b_kath"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                                sum[5] += Convert.ToDecimal(tempRow[5]);
                                sum[6] += Convert.ToDecimal(tempRow[6]);
                                sum[7] += Convert.ToDecimal(tempRow[7]);
                                sum[8] += Convert.ToDecimal(tempRow[8]);
                                sum[9] += Convert.ToDecimal(tempRow[9]);
                                sum[10] += Convert.ToDecimal(tempRow[10]);
                                sum[11] += Convert.ToDecimal(tempRow[11]);
                                sum[12] += Convert.ToDecimal(tempRow[12]);
                                sum[13] += Convert.ToDecimal(tempRow[13]);
                                sum[14] += Convert.ToDecimal(tempRow[14]);
                                sum[15] += Convert.ToDecimal(tempRow[15]);
                                sum[16] += Convert.ToDecimal(tempRow[16]);
                            }

                            DataRow row2 = (DataRow)table.NewRow();

                            row2[0] = sum[0];
                            row2[1] = sum[1];
                            row2[2] = sum[2];
                            row2[3] = sum[3];
                            row2[4] = sum[4];
                            row2[5] = sum[5];
                            row2[6] = sum[6];
                            row2[7] = sum[7];
                            row2[8] = sum[8];
                            row2[9] = sum[9];
                            row2[10] = sum[10];
                            row2[11] = sum[11];
                            row2[12] = sum[12];
                            row2[13] = sum[13];
                            row2[14] = sum[14];
                            row2[15] = sum[15];
                            row2[16] = sum[16];
                            table.Rows.Clear();
                            table.Rows.Add(row2);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView4.DataSource = table;
                            dataGridView4.Columns["date"].DisplayIndex = 0;
                            dataGridView4.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView4.Columns[0].HeaderText = "Συνολική Παροχή (m³/d)";
                            dataGridView4.Columns[1].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView4.Columns[2].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView4.Columns[3].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView4.Columns[4].HeaderText = "Ταχύτητα Εξ. (m²/d)";
                            dataGridView4.Columns[5].HeaderText = "Φορτίο SS out (Kg/d)";
                            dataGridView4.Columns[6].HeaderText = "BOD Απομακρυνόμενο (Kg/d)";
                            dataGridView4.Columns[7].HeaderText = "BOD Απόδοση (%)";
                            dataGridView4.Columns[8].HeaderText = "SS Απόδοση (%)";
                            dataGridView4.Columns[9].HeaderText = "Φορτιο Ν-ΝΗ3 out (Kg/d)";
                            dataGridView4.Columns[10].HeaderText = "Φορτιο Ν-ΝΟ3 out (Kg/d)";
                            dataGridView4.Columns[11].HeaderText = "N Νιτροποίησης (Kg/d)";
                            dataGridView4.Columns[12].HeaderText = "N Απονιτροποίησης (Kg/d)";
                            dataGridView4.Columns[13].HeaderText = "Φορτίο Ptot out (Kg/d)";
                            dataGridView4.Columns[14].HeaderText = "Κατανάλωση Fe/απομακρυνόμενου P (Kg/Kg)";
                            dataGridView4.Columns[15].HeaderText = "Απόδοση απομ. Ptot β' καθ. (%)";
                            dataGridView4.Columns[16].HeaderText = "Απόδοση απομ. Ptot α' καθ. (%)";

                            dataGridView4.Rows[dataGridView4.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView4.Rows[dataGridView4.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM ol_apod_eel WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "ol_apod_eel");

                            table = ds1.Tables["ol_apod_eel"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                sum[0] += Convert.ToDecimal(tempRow[0]);
                                sum[1] += Convert.ToDecimal(tempRow[1]);
                                sum[2] += Convert.ToDecimal(tempRow[2]);
                                sum[3] += Convert.ToDecimal(tempRow[3]);
                                sum[4] += Convert.ToDecimal(tempRow[4]);
                                sum[5] += Convert.ToDecimal(tempRow[5]);
                            }

                            DataRow row3 = (DataRow)table.NewRow();

                            row3[0] = sum[0];
                            row3[1] = sum[1];
                            row3[2] = sum[2];
                            row3[3] = sum[3];
                            row3[4] = sum[4];
                            row3[5] = sum[5];
                            table.Rows.Clear();
                            table.Rows.Add(row3);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView5.DataSource = table;
                            dataGridView5.Columns["date"].DisplayIndex = 0;
                            dataGridView5.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView5.Columns[0].HeaderText = "Απόδοση απομ. Ptot (%)";
                            dataGridView5.Columns[1].HeaderText = "Απόδοση απομ. N-NH3 (%)";
                            dataGridView5.Columns[2].HeaderText = "Απόδοση απομ. N-NO3 (%)";
                            dataGridView5.Columns[3].HeaderText = "Απόδοση απομ. BOD (%)";
                            dataGridView5.Columns[4].HeaderText = "Απόδοση απομ. SS (%)";
                            dataGridView5.Columns[5].HeaderText = "Απόδοση απομ. Ntot (%)";

                            dataGridView5.Rows[dataGridView5.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView5.Rows[dataGridView5.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM dex_aer WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "dex_aer");

                            table = ds1.Tables["dex_aer"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 32; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row4 = (DataRow)table.NewRow();

                            for (int i = 0; i < 32; i++)
                            {
                                row4[i] = sum[i];
                            }
                            table.Rows.Clear();
                            table.Rows.Add(row4);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView6.DataSource = table;
                            dataGridView6.Columns["date"].DisplayIndex = 0;
                            dataGridView6.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView6.Columns[0].HeaderText = "Χρόνος Παραμονής στον αερ. (h)";
                            dataGridView6.Columns[1].HeaderText = "Ποσοστό Πτητικών (%)";
                            dataGridView6.Columns[2].HeaderText = "Ολικά MLSS αερ. (Kg)";
                            dataGridView6.Columns[3].HeaderText = "Ολικά MLVSS αερ. (Kg)";
                            dataGridView6.Columns[4].HeaderText = "Παροχή αέρα (m³/d)";
                            dataGridView6.Columns[5].HeaderText = "Λόγος αέρα προς παρ. εισ. (m³/m³)";
                            dataGridView6.Columns[6].HeaderText = "BOD απομακρυνόμενο (m³/Kg)";
                            dataGridView6.Columns[7].HeaderText = "Λόγος αέρα προς απομακρυνόμενο BOD (Kg/m³)";
                            dataGridView6.Columns[8].HeaderText = "Απόδοση σε Ο2 (Kg/m³)";
                            dataGridView6.Columns[9].HeaderText = "Παροχή O2 (Kg/d)";
                            dataGridView6.Columns[10].HeaderText = "Λόγος O2 προς απομακρυνόμενο BOD (Kg/Kg)";
                            dataGridView6.Columns[11].HeaderText = "Oγκομετρικό φορτίο BOD (Kg/d*m³)";
                            dataGridView6.Columns[12].HeaderText = "Συνολική Παροχή Αέρα (m³/d)";
                            dataGridView6.Columns[13].HeaderText = "VSS υπολογισμός (Kg/d)";
                            dataGridView6.Columns[14].HeaderText = "Inerts υπολογισμός (Kg/d)";
                            dataGridView6.Columns[15].HeaderText = "TSS υπολογισμός (Kg/d)";
                            dataGridView6.Columns[16].HeaderText = "TSS στη διαδικασία (Kg)";
                            dataGridView6.Columns[17].HeaderText = "MCRT (d)";
                            dataGridView6.Columns[18].HeaderText = "F/M (Kg*d/Kg)";
                            dataGridView6.Columns[19].HeaderText = "Φορτίο Ν-ΝΗ3 εισ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[20].HeaderText = "Φορτίο Ν-ΝΟ3 εισ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[21].HeaderText = "Φορτίο Ν-ΝΗ3 εξ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[22].HeaderText = "Φορτίο Ν-ΝΟ3 εξ. Βιολ. (Kg/d)";
                            dataGridView6.Columns[23].HeaderText = "N Νιτροποίησης (Kg/d)";
                            dataGridView6.Columns[24].HeaderText = "N Απονιτροποίησης (Kg/d)";
                            dataGridView6.Columns[25].HeaderText = "Χρόνος Αερισμού (h)";
                            dataGridView6.Columns[26].HeaderText = "Χρονος Απονιτροποίησης (h)";
                            dataGridView6.Columns[27].HeaderText = "Απόδοση απομ. Ν-ΝΗ3 (%)";
                            dataGridView6.Columns[28].HeaderText = "Απόδοση απομ. Ν-ΝO3 (%)";
                            dataGridView6.Columns[29].HeaderText = "Φορτίο Ptot out (Kg/d)";
                            dataGridView6.Columns[30].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView6.Columns[31].HeaderText = "Απόδοση απομ. Ptot (%)";

                            dataGridView6.Rows[dataGridView6.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView6.Rows[dataGridView6.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM c_vathm WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "c_vathm");

                            table = ds1.Tables["c_vathm"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 3; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row5 = (DataRow)table.NewRow();

                            for (int i = 0; i < 3; i++)
                            {
                                row5[i] = sum[i];
                            }
                            table.Rows.Clear();
                            table.Rows.Add(row5);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView7.DataSource = ds1.Tables["c_vathm"];
                            dataGridView7.Columns["date"].DisplayIndex = 0;
                            dataGridView7.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView7.Columns[0].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView7.Columns[1].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView7.Columns[2].HeaderText = "Απόδοση Φίλτρων (%)";

                            dataGridView7.Rows[dataGridView7.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView7.Rows[dataGridView7.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM propax WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "propax");

                            table = ds1.Tables["propax"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 6; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row6 = (DataRow)table.NewRow();

                            for (int i = 0; i < 6; i++)
                            {
                                row6[i] = sum[i];
                            }
                            table.Rows.Clear();
                            table.Rows.Add(row6);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView8.DataSource = ds1.Tables["propax"];
                            dataGridView8.Columns["date"].DisplayIndex = 0;
                            dataGridView8.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView8.Columns[0].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView8.Columns[1].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView8.Columns[2].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView8.Columns[3].HeaderText = "Φορτίο SS παχυμένης (kg/d)";
                            dataGridView8.Columns[4].HeaderText = "Φορτίο SS στραγγιδίων (kg/d)";
                            dataGridView8.Columns[5].HeaderText = "SS στραγγιδίων (mg/l)";

                            dataGridView8.Rows[dataGridView8.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView8.Rows[dataGridView8.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM mhx_pax WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "mhx_pax");

                            table = ds1.Tables["mhx_pax"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 9; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row7 = (DataRow)table.NewRow();

                            for (int i = 0; i < 9; i++)
                            {
                                row7[i] = sum[i];
                            }
                            table.Rows.Clear();
                            table.Rows.Add(row7);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView9.DataSource = ds1.Tables["mhx_pax"];
                            dataGridView9.Columns["date"].DisplayIndex = 0;
                            dataGridView9.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView9.Columns[0].HeaderText = "Φορτίο WASS (kg/d)";
                            dataGridView9.Columns[1].HeaderText = "Παροχή προς ομογ. (m³/d)";
                            dataGridView9.Columns[2].HeaderText = "Φορτίο SS παχυμ. (kg/d)";
                            dataGridView9.Columns[3].HeaderText = "Φορτίο SS στραγγ. (kg/d)";
                            dataGridView9.Columns[4].HeaderText = "SS στραγγ. (mg/l)";
                            dataGridView9.Columns[5].HeaderText = "Ειδική καταν. ΠΗΛ (kg/Kg)";
                            dataGridView9.Columns[6].HeaderText = "Λόγος Πάχυνσης (decimal)";
                            dataGridView9.Columns[7].HeaderText = "Φορτίο ομογεν. (kg/d)";
                            dataGridView9.Columns[8].HeaderText = "SS παχυμ. (mg/l)";

                            dataGridView9.Rows[dataGridView9.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView9.Rows[dataGridView9.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM xwneusi WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "xwneusi");

                            table = ds1.Tables["xwneusi"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 12; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row8 = (DataRow)table.NewRow();

                            for (int i = 0; i < 12; i++)
                            {
                                row8[i] = sum[i];
                            }
                            table.Rows.Clear();
                            table.Rows.Add(row8);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView10.DataSource = ds1.Tables["xwneusi"];
                            dataGridView10.Columns["date"].DisplayIndex = 0;
                            dataGridView10.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView10.Columns[0].HeaderText = "Χρόνος Παραμονής (d)";
                            dataGridView10.Columns[1].HeaderText = "Φορτίο SS χώνευσης (kg/d)";
                            dataGridView10.Columns[2].HeaderText = "Φορτίο VSS χώνευσης (kg/d)";
                            dataGridView10.Columns[3].HeaderText = "Ποσοστό Πτητικών in (%)";
                            dataGridView10.Columns[4].HeaderText = "Φόρτιση Πτητικών (kg/d*m³)";
                            dataGridView10.Columns[5].HeaderText = "Φορτίο SS fix (kg/d)";
                            dataGridView10.Columns[6].HeaderText = "Φορτίο SS εξόδου (kg/d)";
                            dataGridView10.Columns[7].HeaderText = "Φορτίο VSS εξόδου (kg/d)";
                            dataGridView10.Columns[8].HeaderText = "Ποσοστό Πτητικών out (%)";
                            dataGridView10.Columns[9].HeaderText = "Παραγωγή Βιοαερίου (kg/d)";
                            dataGridView10.Columns[10].HeaderText = "Απόδοση (%)";
                            dataGridView10.Columns[11].HeaderText = "Λόγος VA/A (decimal)";

                            dataGridView10.Rows[dataGridView10.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView10.Rows[dataGridView10.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM metapax WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "metapax");

                            table = ds1.Tables["metapax"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 6; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row9 = (DataRow)table.NewRow();

                            for (int i = 0; i < 6; i++)
                            {
                                row9[i] = sum[i];
                            }
                            table.Rows.Clear();
                            table.Rows.Add(row9);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView11.DataSource = ds1.Tables["metapax"];
                            dataGridView11.Columns["date"].DisplayIndex = 0;
                            dataGridView11.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView11.Columns[0].HeaderText = "Χρόνος Παραμονής (h)";
                            dataGridView11.Columns[1].HeaderText = "Υδραυλική Φόρτιση (m/d)";
                            dataGridView11.Columns[2].HeaderText = "Φόρτιση Στερεών (kg/d*m²)";
                            dataGridView11.Columns[3].HeaderText = "Φορτίο SS παχυμένης (kg/d)";
                            dataGridView11.Columns[4].HeaderText = "Φορτίο SS στραγγιδίων (kg/d)";
                            dataGridView11.Columns[5].HeaderText = "SS στραγγιδίων (mg/l)";

                            dataGridView11.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView11.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;

                            com = "SELECT * FROM afudatwsi WHERE date>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND date<='" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";

                            adpt = new MySqlDataAdapter(com, connection1);
                            adpt.Fill(ds1, "afudatwsi");

                            table = ds1.Tables["afudatwsi"];

                            tempRow = null;

                            foreach (DataRow tempRow_Variable in table.Rows)
                            {
                                tempRow = tempRow_Variable;
                                for (int i = 0; i < 3; i++)
                                {
                                    sum[i] += Convert.ToDecimal(tempRow[i]);
                                }
                            }

                            DataRow row10 = (DataRow)table.NewRow();

                            for (int i = 0; i < 3; i++)
                            {
                                row10[i] = sum[i];
                            }
                            table.Rows.Clear();
                            table.Rows.Add(row10);

                            Array.Clear(sum, 0, sum.Length);

                            dataGridView12.DataSource = ds1.Tables["afudatwsi"];
                            dataGridView12.Columns["date"].DisplayIndex = 0;
                            dataGridView12.Columns["date"].HeaderText = "Ημερομηνία";
                            dataGridView12.Columns[0].HeaderText = "Φορτίο SS αφυδατωμένης (kg/d)";
                            dataGridView12.Columns[1].HeaderText = "Φορτίο SS στραγγιδίων (kg/d)";
                            dataGridView12.Columns[2].HeaderText = "Eιδική καταν. ΠΗΛ (kg/Kg)";

                            dataGridView12.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView12.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                        }
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message, "Σφάλμα επικοινωνίας με διακομιστή", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

        }

        private void dataGridView1_MouseHover(object sender, EventArgs e)
        {
            dataGridView1.Focus();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.Show();
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Form3 frm = new Form3();
            frm.Show();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        public void calculations(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(ClassLibrary.Class1.sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();

            string appPath = Path.GetDirectoryName(Application.ExecutablePath) + @"\tags.txt";
            //string appPath = @"c:\Users\" + Environment.UserName + @"\Desktop\sql.txt";
            List<string> lines = new List<string>();

            using (StreamReader r = new StreamReader(appPath, Encoding.Default))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }
            List<string> tags = new List<string>();
            foreach (string s in lines)
            {
                tags.Add(Convert.ToString(ClassLibrary.Class1.integral(s.Trim(), date)));
            }

            command.CommandText = "insert into calculations(A6_FLOW,A3_FLOW,A5_FLOW,PAROXH_A_ILYOS_1,PAROXH_A_ILYOS_3,PAROXH_PROP_1,PAROXH_PROP_2,PERISIA_PAROXI,25IF02,FLOW_OMOG,FLOW,MLSS_A_ILYOS_1,MLSS_A_ILYOS_3,MLSS_1,MLSS_2,MLSS_3,MLSS_4,MLSS_5,PAROXI_EISODOU,SUM,MLSS,MLSS_PERIS_NEW,IDO_1,IDO_2,IDO_3,IDO_4,IDO_05,MLSS_DEC,BLOWERS_POWER,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13, @p14, @p15, @p16, @p17, @p18, @p19, @p20, @p21, @p22, @p23, @p24, @p25, @p26, @p27, @p28, @p29, @p30)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", tags[0]);
            command.Parameters.AddWithValue("@p2", tags[1]);
            command.Parameters.AddWithValue("@p3", tags[2]);
            command.Parameters.AddWithValue("@p4", tags[3]);
            command.Parameters.AddWithValue("@p5", tags[4]);
            command.Parameters.AddWithValue("@p6", tags[5]);
            command.Parameters.AddWithValue("@p7", tags[6]);
            command.Parameters.AddWithValue("@p8", tags[7]);
            command.Parameters.AddWithValue("@p9", tags[8]);
            command.Parameters.AddWithValue("@p10", tags[9]);
            command.Parameters.AddWithValue("@p11", tags[10]);
            command.Parameters.AddWithValue("@p12", tags[11]);
            command.Parameters.AddWithValue("@p13", tags[12]);
            command.Parameters.AddWithValue("@p14", tags[13]);
            command.Parameters.AddWithValue("@p15", tags[14]);
            command.Parameters.AddWithValue("@p16", tags[15]);
            command.Parameters.AddWithValue("@p17", tags[16]);
            command.Parameters.AddWithValue("@p18", tags[17]);
            command.Parameters.AddWithValue("@p19", tags[18]);
            command.Parameters.AddWithValue("@p20", tags[19]);
            command.Parameters.AddWithValue("@p21", tags[20]);
            command.Parameters.AddWithValue("@p22", tags[21]);
            command.Parameters.AddWithValue("@p23", tags[22]);
            command.Parameters.AddWithValue("@p24", tags[23]);
            command.Parameters.AddWithValue("@p25", tags[24]);
            command.Parameters.AddWithValue("@p26", tags[25]);
            command.Parameters.AddWithValue("@p27", tags[26]);
            command.Parameters.AddWithValue("@p28", tags[27]);
            command.Parameters.AddWithValue("@p29", tags[28]);
            command.Parameters.AddWithValue("@p30", date);
            command.ExecuteNonQuery();

            ClassLibrary.Class1.eisodos_eel(date);
            ClassLibrary.Class1.a_kath(date);
            ClassLibrary.Class1.dex_aer(date);
            ClassLibrary.Class1.b_kath(date);
            ClassLibrary.Class1.c_vathm(date);
            ClassLibrary.Class1.ol_apod_eel(date);
            ClassLibrary.Class1.propax(date);
            ClassLibrary.Class1.mhx_pax(date);
            ClassLibrary.Class1.xwneusi(date);
            ClassLibrary.Class1.metapax(date);
            ClassLibrary.Class1.afudat(date);

            MessageBox.Show("Οι πράξεις ολοκληρώθηκαν");

            connection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //try
            //{
                calculations(DateTime.Today.AddDays(-1));
            //}
            //catch (Exception exc)
            //{
            //    MessageBox.Show(exc.Message, "Σφάλμα επικοινωνίας με διακομιστή", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void label47_Click(object sender, EventArgs e)
        {

        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {

        }

        private void label48_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label50_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void label57_Click(object sender, EventArgs e)
        {

        }

        private void label56_Click(object sender, EventArgs e)
        {

        }

        private void label65_Click(object sender, EventArgs e)
        {

        }

        private void label64_Click(object sender, EventArgs e)
        {

        }

        private void label69_Click(object sender, EventArgs e)
        {

        }

        private void label68_Click(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox35_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox34_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox33_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox36_TextChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void label84_Click(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Location = new Point(123, 74);
            dateTimePicker2.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            tabControl1.Visible = false;
            tabControl2.Visible = true;
            dataGridView1.Visible = false;
            comboBox1.Visible = false;
            checkBox1.Visible = true;
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Location = new Point(305, 74);
            dateTimePicker2.Visible = false;
            label1.Visible = false;
            label2.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            tabControl1.Visible = true;
            tabControl2.Visible = false;
            dataGridView1.Visible = false;
            comboBox1.Visible = false;
            checkBox1.Visible = false;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //Bitmap bm = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);
            //dataGridView1.DrawToBitmap(bm, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));
            //e.Graphics.DrawImage(bm, 0, 0);


            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView2.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView2.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView2.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΕΙΣΟΔΟΣ ΕΕΛ", new Font(dataGridView2.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView2.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView1.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView1.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView1.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView2.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            //Drawing Cells Borders 
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (tabControl1.Visible == false)
            {
                printDialog1.Document = printDocument1;
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    printDocument1.DefaultPageSettings.Landscape = true;
                    printDocument1.Print();
                    printDocument2.DefaultPageSettings.Landscape = true;
                    printDocument2.Print();
                    printDocument3.DefaultPageSettings.Landscape = true;
                    printDocument3.Print();
                    printDocument4.DefaultPageSettings.Landscape = true;
                    printDocument4.Print();
                    printDocument5.DefaultPageSettings.Landscape = true;
                    printDocument5.Print();
                    printDocument6.DefaultPageSettings.Landscape = true;
                    printDocument6.Print();
                    printDocument7.DefaultPageSettings.Landscape = true;
                    printDocument7.Print();
                    printDocument8.DefaultPageSettings.Landscape = true;
                    printDocument8.Print();
                    printDocument9.DefaultPageSettings.Landscape = true;
                    printDocument9.Print();
                    printDocument10.DefaultPageSettings.Landscape = true;
                    printDocument10.Print();
                    printDocument11.DefaultPageSettings.Landscape = true;
                    printDocument11.Print();
                }
            }
            else
            {
                 printDialog1.Document = printDocument12;
                 if (printDialog1.ShowDialog() == DialogResult.OK)
                 {
                     printDocument12.Print();
                     printDocument13.Print();
                     printDocument14.Print();
                     printDocument15.Print();
                     printDocument16.Print();
                     printDocument17.Print();
                     printDocument18.Print();
                     printDocument19.Print();
                     printDocument20.Print();
                     printDocument21.Print();
                     printDocument22.Print();
                 }
            }
        }

        private void TAGToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Location = new Point(123, 74);
            dateTimePicker2.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            tabControl1.Visible = false;
            tabControl2.Visible = false;
            dataGridView1.Visible = true;
            comboBox1.Visible = true;
            checkBox1.Visible = false;
        }

        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {

            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView2.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument2_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView3.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView3.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView3.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("Α' ΚΑΘΙΖΗΣΗ", new Font(dataGridView3.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView3.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView3.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView3.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView3.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView3.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument2_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView3.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument3_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView4.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument3_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView4.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView4.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView4.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("B' ΚΑΘΙΖΗΣΗ", new Font(dataGridView4.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView4.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView4.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView4.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView4.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView4.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iCellHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument4_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView5.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument4_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView5.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView5.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView5.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΟΛΙΚΕΣ ΑΠΟΔΟΣΕΙΣ ΕΕΛ", new Font(dataGridView5.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView5.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView5.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView5.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView5.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView5.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument5_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView6.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView6.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView6.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΔΕΞΑΜΕΝΗ ΑΕΡΙΣΜΟΥ", new Font(dataGridView6.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView6.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView6.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView6.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView6.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView6.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iCellHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument5_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView6.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument6_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView7.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument6_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView7.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 40;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView7.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView7.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("Γ' ΕΠΕΞΕΡΓΑΣΙΑ", new Font(dataGridView7.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView7.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView7.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView7.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView7.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView7.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iCellHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument7_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView8.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument7_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView8.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView8.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView8.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΠΡΟΠΑΧΥΝΣΗ", new Font(dataGridView8.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView8.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView8.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView8.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView8.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView8.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument8_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView9.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument8_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView9.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView9.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView9.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΜΗΧΑΝΙΚΗ ΠΑΧΥΝΣΗ", new Font(dataGridView9.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView9.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView9.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView9.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView9.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView9.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument9_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView10.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument9_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView10.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView10.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView10.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΧΩΝΕΥΣΗ", new Font(dataGridView10.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView10.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView10.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView10.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView10.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView10.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument10_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView11.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument10_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView11.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 40;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView11.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView11.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΜΕΤΑΠΑΧΥΝΣΗ", new Font(dataGridView11.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView11.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView11.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView11.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView11.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView11.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iCellHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument11_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView12.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument11_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView12.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dataGridView12.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView12.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString("ΑΦΥΔΑΤΩΣΗ", new Font(dataGridView12.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Customer Summary", new Font(dataGridView12.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            if (dateTimePicker2.Visible == true)
                            {
                                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                            }
                            else
                            {
                                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                            }
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dataGridView12.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView12.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ΜΕΤΡΗΣΕΙΣ", new Font(new Font(dataGridView12.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView12.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            e.Graphics.FillRectangle(new SolidBrush(Cel.InheritedStyle.BackColor),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument12_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage1.Width, this.tabPage1.Height);
            tabPage1.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage1.Width, this.tabPage1.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument13_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage2.Width, this.tabPage2.Height);
            tabPage2.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage2.Width, this.tabPage2.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument14_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage3.Width, this.tabPage3.Height);
            tabPage3.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage3.Width, this.tabPage3.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument15_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage4.Width, this.tabPage4.Height);
            tabPage4.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage4.Width, this.tabPage4.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument16_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage5.Width, this.tabPage5.Height);
            tabPage5.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage5.Width, this.tabPage5.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument17_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage6.Width, this.tabPage6.Height);
            tabPage6.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage6.Width, this.tabPage6.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument18_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage7.Width, this.tabPage7.Height);
            tabPage7.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage7.Width, this.tabPage7.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument19_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage8.Width, this.tabPage8.Height);
            tabPage8.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage8.Width, this.tabPage8.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument20_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage9.Width, this.tabPage9.Height);
            tabPage9.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage9.Width, this.tabPage9.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument21_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage10.Width, this.tabPage10.Height);
            tabPage10.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage10.Width, this.tabPage10.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void printDocument22_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String strDate = "";
            if (dateTimePicker2.Visible == true)
            {
                strDate = "ΑΠΟ " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + " ΕΩΣ " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
            }
            else
            {
                strDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");
            }
            //Draw Date
            e.Graphics.DrawString(strDate, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 0, 0);

            Bitmap bm = new Bitmap(this.tabPage11.Width, this.tabPage11.Height);
            tabPage11.DrawToBitmap(bm, new Rectangle(0, 0, this.tabPage11.Width, this.tabPage11.Height));
            e.Graphics.DrawImage(bm, 0, 40);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void ToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            if (tabControl1.Visible == false)
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Worksheet xlWorkSheet3;
                Excel.Worksheet xlWorkSheet4;
                Excel.Worksheet xlWorkSheet5;
                Excel.Worksheet xlWorkSheet6;
                Excel.Worksheet xlWorkSheet7;
                Excel.Worksheet xlWorkSheet8;
                Excel.Worksheet xlWorkSheet9;
                Excel.Worksheet xlWorkSheet10;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet3 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet4 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet5 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet6 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet7 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet8 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet9 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet10 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                int i = 0;
                int j = 0;

                xlWorkSheet.Name = "ΕΙΣΟΔΟΣ ΕΕΛ";
                //xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                //Excel.Range headerRange = xlWorkSheet.get_Range("A1","V1");
                //headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                for (i = 1; i < dataGridView2.Columns.Count + 1; i++)
                {
                    xlWorkSheet.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                }
                xlWorkSheet.Columns.AutoFit();
                for (i = 0; i <= dataGridView2.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView2[j, i];
                        xlWorkSheet.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if(checkBox1.Checked)
                {
                    xlWorkSheet.Cells[2, 6] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                Excel.Range formatRange;
                formatRange = xlWorkSheet.get_Range("a"+ i, "e"+ i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;
                
                Excel.Worksheet xlWorkSheet1;
                xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                xlWorkSheet1.Name = "Α' ΚΑΘΙΖΗΣΗ";
                //xlWorkSheet1.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView3.Columns.Count + 1; i++)
                {
                    xlWorkSheet1.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
                }
                xlWorkSheet1.Columns.AutoFit();
                for (i = 0; i <= dataGridView3.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView3.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView3[j, i];
                        xlWorkSheet1.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet1.Cells[2, 9] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet1.get_Range("a" + i, "h" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                Excel.Worksheet xlWorkSheet2;
                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
                xlWorkSheet2.Name = "Β' ΚΑΘΙΖΗΣΗ";
                //xlWorkSheet2.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView4.Columns.Count + 1; i++)
                {
                    xlWorkSheet2.Cells[1, i] = dataGridView4.Columns[i - 1].HeaderText;
                }
                xlWorkSheet2.Columns.AutoFit();
                for (i = 0; i <= dataGridView4.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView4.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView4[j, i];
                        xlWorkSheet2.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet2.Cells[2, 18] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet2.get_Range("a" + i, "q" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet3 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);
                xlWorkSheet3.Name = "ΟΛΙΚΕΣ ΑΠΟΔΟΣΕΙΣ ΕΕΛ";
                //xlWorkSheet3.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView5.Columns.Count + 1; i++)
                {
                    xlWorkSheet3.Cells[1, i] = dataGridView5.Columns[i - 1].HeaderText;
                }
                xlWorkSheet3.Columns.AutoFit();
                for (i = 0; i <= dataGridView5.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView5.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView5[j, i];
                        xlWorkSheet3.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet3.Cells[2, 7] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet3.get_Range("a" + i, "f" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet4 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);
                xlWorkSheet4.Name = "ΔΕΞΑΜΕΝΗ ΑΕΡΙΣΜΟΥ";
                //xlWorkSheet4.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView6.Columns.Count + 1; i++)
                {
                    xlWorkSheet4.Cells[1, i] = dataGridView6.Columns[i - 1].HeaderText;
                }
                xlWorkSheet4.Columns.AutoFit();
                for (i = 0; i <= dataGridView6.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView6.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView6[j, i];
                        xlWorkSheet4.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet4.Cells[2, 33] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet4.get_Range("a" + i, "af" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet5 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);
                xlWorkSheet5.Name = "Γ' ΕΠΕΞΕΡΓΑΣΙΑ";
                //xlWorkSheet5.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView7.Columns.Count + 1; i++)
                {
                    xlWorkSheet5.Cells[1, i] = dataGridView7.Columns[i - 1].HeaderText;
                }
                xlWorkSheet5.Columns.AutoFit();
                for (i = 0; i <= dataGridView7.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView7.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView7[j, i];
                        xlWorkSheet5.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet5.Cells[2, 4] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet5.get_Range("a" + i, "c" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet6 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(7);
                xlWorkSheet6.Name = "ΠΡΟΠΑΧΥΝΣΗ";
                //xlWorkSheet6.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView8.Columns.Count + 1; i++)
                {
                    xlWorkSheet6.Cells[1, i] = dataGridView8.Columns[i - 1].HeaderText;
                }
                xlWorkSheet6.Columns.AutoFit();
                for (i = 0; i <= dataGridView8.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView8.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView8[j, i];
                        xlWorkSheet6.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet6.Cells[2, 7] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet6.get_Range("a" + i, "f" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet7 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(8);
                xlWorkSheet7.Name = "ΜΗΧΑΝΙΚΗ ΠΑΧΥΝΣΗ";
                //xlWorkSheet7.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView9.Columns.Count + 1; i++)
                {
                    xlWorkSheet7.Cells[1, i] = dataGridView9.Columns[i - 1].HeaderText;
                }
                xlWorkSheet7.Columns.AutoFit();
                for (i = 0; i <= dataGridView9.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView9.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView9[j, i];
                        xlWorkSheet7.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet7.Cells[2, 10] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet7.get_Range("a" + i, "i" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet8 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(9);
                xlWorkSheet8.Name = "ΧΩΝΕΥΣΗ";
                //xlWorkSheet8.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView10.Columns.Count + 1; i++)
                {
                    xlWorkSheet8.Cells[1, i] = dataGridView10.Columns[i - 1].HeaderText;
                }
                xlWorkSheet8.Columns.AutoFit();
                for (i = 0; i <= dataGridView10.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView10.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView10[j, i];
                        xlWorkSheet8.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet8.Cells[2, 13] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet8.get_Range("a" + i, "l" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet9 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(10);
                xlWorkSheet9.Name = "ΜΕΤΑΠΑΧΥΝΣΗ";
                //xlWorkSheet9.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView11.Columns.Count + 1; i++)
                {
                    xlWorkSheet9.Cells[1, i] = dataGridView11.Columns[i - 1].HeaderText;
                }
                xlWorkSheet9.Columns.AutoFit();
                for (i = 0; i <= dataGridView11.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView11.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView11[j, i];
                        xlWorkSheet9.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet9.Cells[2, 7] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet9.get_Range("a" + i, "f" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                xlWorkSheet10 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(11);
                xlWorkSheet10.Name = "ΑΦΥΔΑΤΩΣΗ";
                //xlWorkSheet10.Cells[1, 1].EntireRow.Font.Bold = true;
                for (i = 1; i < dataGridView12.Columns.Count + 1; i++)
                {
                    xlWorkSheet10.Cells[1, i] = dataGridView12.Columns[i - 1].HeaderText;
                }
                xlWorkSheet10.Columns.AutoFit();
                for (i = 0; i <= dataGridView12.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView12.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView12[j, i];
                        xlWorkSheet10.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                if (checkBox1.Checked)
                {
                    xlWorkSheet10.Cells[2, 4] = dateTimePicker1.Value.ToString("dd-MM-yyyy") + " - " + dateTimePicker2.Value.ToString("dd-MM-yyyy");
                }
                formatRange = xlWorkSheet10.get_Range("a" + i, "c" + i);
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                formatRange.Font.Bold = true;

                saveFileDialog1.InitialDirectory = @"C:\";
                saveFileDialog1.ShowDialog();
                string path = saveFileDialog1.FileName;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkBook);
                        releaseObject(xlApp);

                        MessageBox.Show("Δημιουργήθηκε έγγραφο excel επιτυχώς");
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Σφάλμα αποθήκευσης", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Worksheet xlWorkSheet3;
                Excel.Worksheet xlWorkSheet4;
                Excel.Worksheet xlWorkSheet5;
                Excel.Worksheet xlWorkSheet6;
                Excel.Worksheet xlWorkSheet7;
                Excel.Worksheet xlWorkSheet8;
                Excel.Worksheet xlWorkSheet9;
                Excel.Worksheet xlWorkSheet10;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet3 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet4 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet5 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet6 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet7 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet8 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet9 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet10 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                int i = 0;
                int j = 0;

                xlWorkSheet.Name = "ΕΙΣΟΔΟΣ ΕΕΛ";
                //xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                //Excel.Range headerRange = xlWorkSheet.get_Range("A1","V1");
                //headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Cells[1, 1] = "Στραγγίδια (m³/d)";
                xlWorkSheet.Cells[1, 2] = "Προπάχυνση (m³/d)";
                xlWorkSheet.Cells[1, 3] = "Μηχανική Πάχυνση (m³/d)";
                xlWorkSheet.Cells[1, 4] = "Μεταπάχυνση (m³/d)";
                xlWorkSheet.Cells[1, 5] = "Αφυδάτωση (m³/d)";
                xlWorkSheet.Cells[1, 6] = "Ημερομηνία";
                
                xlWorkSheet.Columns.AutoFit();

                for (i = 1; i <= 5; i++)
                {
                    TextBox tbox = (TextBox)tabPage1.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet.Cells[2, i] = tbox.Text;
                }
                //j = 1;
                //foreach (var tb in tabPage1.Controls.OfType<TextBox>())
                //{
                //    xlWorkSheet.Cells[2, j] = tb.Text;
                //    j++;
                //}
                xlWorkSheet.Cells[2, 6] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                Excel.Worksheet xlWorkSheet1;
                xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                xlWorkSheet1.Name = "Α' ΚΑΘΙΖΗΣΗ";
                //xlWorkSheet1.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet1.Cells[1, 1] = "Χρόνος Παραμονής (h)";
                xlWorkSheet1.Cells[1, 2] = "Παροχή Α' Ιλύς (m³/d)";
                xlWorkSheet1.Cells[1, 3] = "SS a (mg/l)";
                xlWorkSheet1.Cells[1, 4] = "Φορτίο SS a (Kg/d)";
                xlWorkSheet1.Cells[1, 5] = "Φορτίο Ptot in (Kg/d)";
                xlWorkSheet1.Cells[1, 6] = "Φορτίο Ptot out (Kg/d)";
                xlWorkSheet1.Cells[1, 7] = "Κροκιδωτικό FE (Kg/d)";
                xlWorkSheet1.Cells[1, 8] = "Κατανάλωση FE/Ptot Απομ. (Kg/Kg)";
                xlWorkSheet1.Cells[1, 9] = "Ημερομηνία";

                xlWorkSheet1.Columns.AutoFit();

                j = 1;
                for (i = 6; i <= 13; i++)
                {
                    TextBox tbox = (TextBox)tabPage2.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet1.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet1.Cells[2, 9] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                Excel.Worksheet xlWorkSheet2;
                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
                xlWorkSheet2.Name = "Β' ΚΑΘΙΖΗΣΗ";
                //xlWorkSheet2.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet2.Cells[1, 1] = "Συνολική Παροχή (m³/d)";
                xlWorkSheet2.Cells[1, 2] = "Υδραυλική Φόρτιση (m/d)";
                xlWorkSheet2.Cells[1, 3] = "Φόρτιση Στερεών (kg/d*m²)";
                xlWorkSheet2.Cells[1, 4] = "Χρόνος Παραμονής (h)";
                xlWorkSheet2.Cells[1, 5] = "Ταχύτητα Εξ. (m²/d)";
                xlWorkSheet2.Cells[1, 6] = "Φορτίο SS out (Kg/d)";
                xlWorkSheet2.Cells[1, 7] = "BOD Απομακρυνόμενο (Kg/d)";
                xlWorkSheet2.Cells[1, 8] = "BOD Απόδοση (%)";
                xlWorkSheet2.Cells[1, 9] = "SS Απόδοση (%)";
                xlWorkSheet2.Cells[1, 10] = "Φορτιο Ν-ΝΗ3 out (Kg/d)";
                xlWorkSheet2.Cells[1, 11] = "Φορτιο Ν-ΝΟ3 out (Kg/d)";
                xlWorkSheet2.Cells[1, 12] = "N Νιτροποίησης (Kg/d)";
                xlWorkSheet2.Cells[1, 13] = "N Απονιτροποίησης (Kg/d)";
                xlWorkSheet2.Cells[1, 14] = "Φορτίο Ptot out (Kg/d)";
                xlWorkSheet2.Cells[1, 15] = "Κατανάλωση Fe/απομακρυνόμενου P (Kg/Kg)";
                xlWorkSheet2.Cells[1, 16] = "Απόδοση απομ. Ptot β' καθ. (%)";
                xlWorkSheet2.Cells[1, 17] = "Απόδοση απομ. Ptot α' καθ. (%)";
                xlWorkSheet2.Cells[1, 18] = "Ημερομηνία";

                xlWorkSheet2.Columns.AutoFit();

                j = 1;
                for (i = 14; i <= 30; i++)
                {
                    TextBox tbox = (TextBox)tabPage3.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet2.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet2.Cells[2, 18] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet3 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);
                xlWorkSheet3.Name = "ΟΛΙΚΕΣ ΑΠΟΔΟΣΕΙΣ ΕΕΛ";
                //xlWorkSheet3.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet3.Cells[1, 1] = "Απόδοση απομ. Ptot (%)";
                xlWorkSheet3.Cells[1, 2] = "Απόδοση απομ. N-NH3 (%)";
                xlWorkSheet3.Cells[1, 3] = "Απόδοση απομ. N-NO3 (%)";
                xlWorkSheet3.Cells[1, 4] = "Απόδοση απομ. BOD (%)";
                xlWorkSheet3.Cells[1, 5] = "Απόδοση απομ. SS (%)";
                xlWorkSheet3.Cells[1, 6] = "Απόδοση απομ. Ntot (%)";
                xlWorkSheet3.Cells[1, 7] = "Ημερομηνία";

                xlWorkSheet3.Columns.AutoFit();

                j = 1;
                for (i = 31; i <= 36; i++)
                {
                    TextBox tbox = (TextBox)tabPage4.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet3.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet3.Cells[2, 7] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet4 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);
                xlWorkSheet4.Name = "ΔΕΞΑΜΕΝΗ ΑΕΡΙΣΜΟΥ";
                //xlWorkSheet4.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet4.Cells[1, 1] = "Χρόνος Παραμονής στον αερ. (h)";
                xlWorkSheet4.Cells[1, 2] = "Ποσοστό Πτητικών (%)";
                xlWorkSheet4.Cells[1, 3] = "Ολικά MLSS αερ. (Kg)";
                xlWorkSheet4.Cells[1, 4] = "Ολικά MLVSS αερ. (Kg)";
                xlWorkSheet4.Cells[1, 5] = "Παροχή αέρα (m³/d)";
                xlWorkSheet4.Cells[1, 6] = "Λόγος αέρα προς παρ. εισ. (m³/m³)";
                xlWorkSheet4.Cells[1, 7] = "BOD απομακρυνόμενο (m³/Kg)";
                xlWorkSheet4.Cells[1, 8] = "Λόγος αέρα προς απομακρυνόμενο BOD (Kg/m³)";
                xlWorkSheet4.Cells[1, 9] = "Απόδοση σε Ο2 (Kg/m³)";
                xlWorkSheet4.Cells[1, 10] = "Παροχή O2 (Kg/d)";
                xlWorkSheet4.Cells[1, 11] = "Λόγος O2 προς απομακρυνόμενο BOD (Kg/Kg)";
                xlWorkSheet4.Cells[1, 12] = "Oγκομετρικό φορτίο BOD (Kg/d*m³)";
                xlWorkSheet4.Cells[1, 13] = "Συνολική Παροχή Αέρα (m³/d)";
                xlWorkSheet4.Cells[1, 14] = "VSS υπολογισμός (Kg/d)";
                xlWorkSheet4.Cells[1, 15] = "Inerts υπολογισμός (Kg/d)";
                xlWorkSheet4.Cells[1, 16] = "TSS υπολογισμός (Kg/d)";
                xlWorkSheet4.Cells[1, 17] = "TSS στη διαδικασία (Kg)";
                xlWorkSheet4.Cells[1, 18] = "MCRT (d)";
                xlWorkSheet4.Cells[1, 19] = "F/M (Kg*d/Kg)";
                xlWorkSheet4.Cells[1, 20] = "Φορτίο Ν-ΝΗ3 εισ. Βιολ. (Kg/d)";
                xlWorkSheet4.Cells[1, 21] = "Φορτίο Ν-ΝΟ3 εισ. Βιολ. (Kg/d)";
                xlWorkSheet4.Cells[1, 22] = "Φορτίο Ν-ΝΗ3 εξ. Βιολ. (Kg/d)";
                xlWorkSheet4.Cells[1, 23] = "Φορτίο Ν-ΝΟ3 εξ. Βιολ. (Kg/d)";
                xlWorkSheet4.Cells[1, 24] = "N Νιτροποίησης (Kg/d)";
                xlWorkSheet4.Cells[1, 25] = "N Απονιτροποίησης (Kg/d)";
                xlWorkSheet4.Cells[1, 26] = "Χρόνος Αερισμού (h)";
                xlWorkSheet4.Cells[1, 27] = "Χρονος Απονιτροποίησης (h)";
                xlWorkSheet4.Cells[1, 28] = "Απόδοση απομ. Ν-ΝΗ3 (%)";
                xlWorkSheet4.Cells[1, 29] = "Απόδοση απομ. Ν-ΝO3 (%)";
                xlWorkSheet4.Cells[1, 30] = "Φορτίο Ptot out (Kg/d)";
                xlWorkSheet4.Cells[1, 31] = "Χρόνος Παραμονής (h)";
                xlWorkSheet4.Cells[1, 32] = "Απόδοση απομ. Ptot (%)";
                xlWorkSheet4.Cells[1, 33] = "Ημερομηνία";

                xlWorkSheet4.Columns.AutoFit();

                j = 1;
                for (i = 37; i <= 68; i++)
                {
                    TextBox tbox = (TextBox)tabPage5.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet4.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet4.Cells[2, 33] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet5 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);
                xlWorkSheet5.Name = "Γ' ΕΠΕΞΕΡΓΑΣΙΑ";
                //xlWorkSheet5.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet5.Cells[1, 1] = "Υδραυλική Φόρτιση (m/d)";
                xlWorkSheet5.Cells[1, 2] = "Φόρτιση Στερεών (kg/d*m²)";
                xlWorkSheet5.Cells[1, 3] = "Απόδοση Φίλτρων (%)";
                xlWorkSheet5.Cells[1, 4] = "Ημερομηνία";

                xlWorkSheet5.Columns.AutoFit();

                j = 1;
                for (i = 69; i <= 71; i++)
                {
                    TextBox tbox = (TextBox)tabPage6.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet5.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet5.Cells[2, 4] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet6 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(7);
                xlWorkSheet6.Name = "ΠΡΟΠΑΧΥΝΣΗ";
                //xlWorkSheet6.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet6.Cells[1, 1] = "Χρόνος Παραμονής (h)";
                xlWorkSheet6.Cells[1, 2] = "Υδραυλική Φόρτιση (m/d)";
                xlWorkSheet6.Cells[1, 3] = "Φόρτιση Στερεών (kg/d*m²)";
                xlWorkSheet6.Cells[1, 4] = "Φορτίο SS παχυμένης (kg/d)";
                xlWorkSheet6.Cells[1, 5] = "Φορτίο SS στραγγιδίων (kg/d)";
                xlWorkSheet6.Cells[1, 6] = "SS στραγγιδίων (mg/l)";
                xlWorkSheet6.Cells[1, 7] = "Ημερομηνία";

                xlWorkSheet6.Columns.AutoFit();

                j = 1;
                for (i = 72; i <= 77; i++)
                {
                    TextBox tbox = (TextBox)tabPage7.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet6.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet6.Cells[2, 7] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet7 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(8);
                xlWorkSheet7.Name = "ΜΗΧΑΝΙΚΗ ΠΑΧΥΝΣΗ";
                //xlWorkSheet7.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet7.Cells[1, 1] = "Φορτίο WASS (kg/d)";
                xlWorkSheet7.Cells[1, 2] = "Παροχή προς ομογ. (m³/d)";
                xlWorkSheet7.Cells[1, 3] = "Φορτίο SS παχυμ. (kg/d)";
                xlWorkSheet7.Cells[1, 4] = "Φορτίο SS στραγγ. (kg/d)";
                xlWorkSheet7.Cells[1, 5] = "SS στραγγ. (mg/l)";
                xlWorkSheet7.Cells[1, 6] = "Ειδική καταν. ΠΗΛ (kg/Kg)";
                xlWorkSheet7.Cells[1, 7] = "Λόγος Πάχυνσης (decimal)";
                xlWorkSheet7.Cells[1, 8] = "Φορτίο ομογεν. (kg/d)";
                xlWorkSheet7.Cells[1, 9] = "SS παχυμ. (mg/l)";
                xlWorkSheet7.Cells[1, 10] = "Ημερομηνία";

                xlWorkSheet7.Columns.AutoFit();

                j = 1;
                for (i = 78; i <= 86; i++)
                {
                    TextBox tbox = (TextBox)tabPage8.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet7.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet7.Cells[2, 10] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet8 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(9);
                xlWorkSheet8.Name = "ΧΩΝΕΥΣΗ";
                //xlWorkSheet8.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet8.Cells[1, 1] = "Χρόνος Παραμονής (d)";
                xlWorkSheet8.Cells[1, 2] = "Φορτίο SS χώνευσης (kg/d)";
                xlWorkSheet8.Cells[1, 3] = "Φορτίο VSS χώνευσης (kg/d)";
                xlWorkSheet8.Cells[1, 4] = "Ποσοστό Πτητικών in (%)";
                xlWorkSheet8.Cells[1, 5] = "Φόρτιση Πτητικών (kg/d*m³)";
                xlWorkSheet8.Cells[1, 6] = "Φορτίο SS fix (kg/d)";
                xlWorkSheet8.Cells[1, 7] = "Φορτίο SS εξόδου (kg/d)";
                xlWorkSheet8.Cells[1, 8] = "Φορτίο VSS εξόδου (kg/d)";
                xlWorkSheet8.Cells[1, 9] = "Ποσοστό Πτητικών out (%)";
                xlWorkSheet8.Cells[1, 10] = "Παραγωγή Βιοαερίου (kg/d)";
                xlWorkSheet8.Cells[1, 11] = "Απόδοση (%)";
                xlWorkSheet8.Cells[1, 12] = "Λόγος VA/A (decimal)";
                xlWorkSheet8.Cells[1, 13] = "Ημερομηνία";

                xlWorkSheet8.Columns.AutoFit();

                j = 1;
                for (i = 87; i <= 98; i++)
                {
                    TextBox tbox = (TextBox)tabPage9.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet8.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet8.Cells[2, 13] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet9 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(10);
                xlWorkSheet9.Name = "ΜΕΤΑΠΑΧΥΝΣΗ";
                //xlWorkSheet9.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet9.Cells[1, 1] = "Χρόνος Παραμονής (h)";
                xlWorkSheet9.Cells[1, 2] = "Υδραυλική Φόρτιση (m/d)";
                xlWorkSheet9.Cells[1, 3] = "Φόρτιση Στερεών (kg/d*m²)";
                xlWorkSheet9.Cells[1, 4] = "Φορτίο SS παχυμένης (kg/d)";
                xlWorkSheet9.Cells[1, 5] = "Φορτίο SS στραγγιδίων (kg/d)";
                xlWorkSheet9.Cells[1, 6] = "SS στραγγιδίων (mg/l)";
                xlWorkSheet9.Cells[1, 7] = "Ημερομηνία";

                xlWorkSheet9.Columns.AutoFit();

                j = 1;
                for (i = 99; i <= 104; i++)
                {
                    TextBox tbox = (TextBox)tabPage11.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet9.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet9.Cells[2, 7] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                xlWorkSheet10 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(11);
                xlWorkSheet10.Name = "ΑΦΥΔΑΤΩΣΗ";
                //xlWorkSheet10.Cells[1, 1].EntireRow.Font.Bold = true;

                xlWorkSheet10.Cells[1, 1] = "Φορτίο SS αφυδατωμένης (kg/d)";
                xlWorkSheet10.Cells[1, 2] = "Φορτίο SS στραγγιδίων (kg/d)";
                xlWorkSheet10.Cells[1, 3] = "Eιδική καταν. ΠΗΛ (kg/Kg)";
                xlWorkSheet10.Cells[1, 4] = "Ημερομηνία";

                xlWorkSheet10.Columns.AutoFit();

                j = 1;
                for (i = 105; i <= 107; i++)
                {
                    TextBox tbox = (TextBox)tabPage10.Controls.Find(string.Format("textBox{0}", i), false).FirstOrDefault();
                    xlWorkSheet10.Cells[2, j] = tbox.Text;
                    j++;
                }
                xlWorkSheet10.Cells[2, 4] = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                saveFileDialog1.InitialDirectory = @"C:\";
                saveFileDialog1.ShowDialog();
                string path = saveFileDialog1.FileName;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkBook);
                        releaseObject(xlApp);

                        MessageBox.Show("Δημιουργήθηκε έγγραφο excel επιτυχώς");
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Σφάλμα αποθήκευσης", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void label171_Click(object sender, EventArgs e)
        {

        }

        private void label180_Click(object sender, EventArgs e)
        {

        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void tabPage11_Click(object sender, EventArgs e)
        {

        }

        private void label218_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_MouseHover_1(object sender, EventArgs e)
        {
            dataGridView1.Focus();
        }

        private void dataGridView3_MouseHover(object sender, EventArgs e)
        {
            dataGridView3.Focus();
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count >= 2)
            {
                dataGridView2.Rows[dataGridView2.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView2.Rows[dataGridView2.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView4.Rows[dataGridView4.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView4.Rows[dataGridView4.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView5.Rows[dataGridView5.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView5.Rows[dataGridView5.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView6.Rows[dataGridView6.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView6.Rows[dataGridView6.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView7.Rows[dataGridView7.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView7.Rows[dataGridView7.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView8.Rows[dataGridView8.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView8.Rows[dataGridView8.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView9.Rows[dataGridView9.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView9.Rows[dataGridView9.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView10.Rows[dataGridView10.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView10.Rows[dataGridView10.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView11.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView11.Rows[dataGridView11.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White;
                dataGridView12.Rows[dataGridView12.Rows.Count - 2].DefaultCellStyle.BackColor = Color.Red;
                dataGridView12.Rows[dataGridView12.Rows.Count - 2].DefaultCellStyle.ForeColor = Color.White; 
            }
        }

        private void ToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            Form4 frm = new Form4();
            frm.Show();
        }
    }
}
