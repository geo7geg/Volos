using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Data;
using MySql.Data.MySqlClient;
using MySql.Data;
using System.IO;
using System.Reflection;



namespace ClassLibrary
{
    public class Class1
    {
        private static string connectionString;

        public static string sqlstringtext()
        {
            string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\sql.txt";
            //string appPath = @"C:\Documents and Settings\Administrator\Desktop\Wincc Statistics\sql.txt";
            List<string> lines = new List<string>();

            using (StreamReader r = new StreamReader(appPath, Encoding.Default))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }
            string sqltext = "";
            foreach (string s in lines)
            {

                sqltext = s.Trim();

            }

            return sqltext;
        }

        public static string winccstringtext()
        {
            string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\wincc.txt";
            //string appPath = @"C:\Documents and Settings\Administrator\Desktop\Wincc Statistics\wincc.txt";
            List<string> lines = new List<string>();

            using (StreamReader r = new StreamReader(appPath, Encoding.Default))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }
            string sqltext = "";
            foreach (string s in lines)
            {

                sqltext = s.Trim();

            }

            return sqltext;
        }

        public static double integral(string tag, DateTime dte)
        {
            connectionString = winccstringtext();

            DataSet ds = new DataSet();

            OleDbConnection connection;
            connection = new OleDbConnection(connectionString);

            OleDbDataAdapter adapter;

            DateTime date = new DateTime();
            DateTime date1 = new DateTime();

            date = dte.AddDays(-1);
            date1 = dte;
            
            string sql = null;
            sql = @"Tag:R,'" + tag + "','" + date.ToString("yyyy-MM-dd 21:00:00.00") + "','" + dte.ToString("yyyy-MM-dd 21:00:00.00") + "'";

            connection.Open();

            adapter = new OleDbDataAdapter(sql, connection);
            adapter.Fill(ds);

            DataTable table = ds.Tables[0];

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

            return result;
        }

        public static void eisodos_eel(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            //connection.Open();

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal PAROXH_A_ILYOS_1 = Convert.ToDecimal(reader["PAROXH_A_ILYOS_1"]);
            decimal PAROXH_A_ILYOS_3 = Convert.ToDecimal(reader["PAROXH_A_ILYOS_3"]);
            decimal PAROXH_PROP_1 = Convert.ToDecimal(reader["PAROXH_PROP_1"]);
            decimal PAROXH_PROP_2 = Convert.ToDecimal(reader["PAROXH_PROP_2"]);
            decimal PERISIA_PAROXI = Convert.ToDecimal(reader["PERISIA_PAROXI"]);
            decimal FE02 = Convert.ToDecimal(reader["25IF02"]);
            decimal FLOW_OMOG = Convert.ToDecimal(reader["FLOW_OMOG"]);
            decimal FLOW = Convert.ToDecimal(reader["FLOW"]);

            reader.Close();

            decimal propax = (PAROXH_A_ILYOS_1 + PAROXH_A_ILYOS_3) - (PAROXH_PROP_1 + PAROXH_PROP_2);
            decimal mhx_pax = (PERISIA_PAROXI + FE02) - (PAROXH_PROP_1 + PAROXH_PROP_2 + FLOW_OMOG);
            decimal metapax = FLOW_OMOG - FLOW;
            decimal straggidia = propax + mhx_pax + metapax;

            command.CommandText = "insert into eisodos_eel(straggidia,propax,mhx_pax,metapax,date) values (@p1, @p2, @p3, @p4, @p5)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", straggidia);
            command.Parameters.AddWithValue("@p2", propax);
            command.Parameters.AddWithValue("@p3", mhx_pax);
            command.Parameters.AddWithValue("@p4", metapax);
            command.Parameters.AddWithValue("@p5", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void a_kath(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            //connection.Open();

            command.CommandText = "SELECT * FROM dexamenes ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal ogkos_akath = Convert.ToDecimal(reader["ogkos_akath"]);

            reader.Close();

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal PAROXI_EISODOU = Convert.ToDecimal(reader["PAROXI_EISODOU"]);
            decimal A6_FLOW = Convert.ToDecimal(reader["A6_FLOW"]);
            decimal MLSS_A_ILYOS_1 = Convert.ToDecimal(reader["MLSS_A_ILYOS_1"]) * 1000;
            decimal MLSS_A_ILYOS_3 = Convert.ToDecimal(reader["MLSS_A_ILYOS_3"]) * 1000;

            reader.Close();

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal Ptot_in = Convert.ToDecimal(reader["Ptot_in"]);
            decimal Ptot_out = Convert.ToDecimal(reader["Ptot_out"]);

            reader.Close();

            decimal xronos_param = (ogkos_akath * 24) / PAROXI_EISODOU;
            decimal ssa = MLSS_A_ILYOS_1 + MLSS_A_ILYOS_3;
            decimal fort_ptot_in = A6_FLOW * Ptot_in;
            decimal fort_ptot_out = A6_FLOW * Ptot_out;

            command.CommandText = "insert into a_kath(xronos_param,ssa,fortio_ptot_in,fortio_ptot_out,date) values (@p1, @p2, @p3, @p4, @p5)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", xronos_param);
            command.Parameters.AddWithValue("@p2", ssa);
            command.Parameters.AddWithValue("@p3", fort_ptot_in);
            command.Parameters.AddWithValue("@p4", fort_ptot_out);
            command.Parameters.AddWithValue("@p5", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void dex_aer(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            //connection.Open();

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal MLSS = Convert.ToDecimal(reader["MLSS"]);
            decimal MLVSS = Convert.ToDecimal(reader["MLVSS"]);
            decimal BOD_viol = Convert.ToDecimal(reader["BOD_viol"]);
            decimal SS_viol = Convert.ToDecimal(reader["SS_viol"]);
            decimal BOD_out = Convert.ToDecimal(reader["BOD_out"]);
            decimal VSS = Convert.ToDecimal(reader["VSS"]);
            decimal N_NH3_eis = Convert.ToDecimal(reader["N_NH3_eis"]);
            decimal N_NH3_ex = Convert.ToDecimal(reader["N_NH3_ex"]);
            decimal N_NO3_eis = Convert.ToDecimal(reader["N_NO3_eis"]);
            decimal N_NO3_ex = Convert.ToDecimal(reader["N_NO3_ex"]);
            decimal Ptot_out_dex_aer = Convert.ToDecimal(reader["Ptot_out_dex_aer"]);
            decimal Ptot_out = Convert.ToDecimal(reader["Ptot_out"]);
            
            reader.Close();

            command.CommandText = "SELECT * FROM dexamenes ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal sunoliko_v_dex_aer = Convert.ToDecimal(reader["sunoliko_v_dex_aer"]);
            decimal ogkos_bkath = Convert.ToDecimal(reader["ogkos_bkath"]);
            decimal ogkos_aerovias = Convert.ToDecimal(reader["ogkos_aerovias"]);
            decimal ogkos_anaer = Convert.ToDecimal(reader["ogkos_anaer"]);
            decimal ogkos_anox = Convert.ToDecimal(reader["ogkos_anox"]);

            reader.Close();

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal BLOWERS_POWER = Convert.ToDecimal(reader["BLOWERS_POWER"]);
            decimal PAROXI_EISODOU = Convert.ToDecimal(reader["PAROXI_EISODOU"]);
            decimal SUM = Convert.ToDecimal(reader["SUM"]);

            reader.Close();

            command.CommandText = "SELECT * FROM a_kath ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal paroxi_a_il = Convert.ToDecimal(reader["paroxi_a_il"]);

            reader.Close();

            decimal var = Convert.ToDecimal(1 / 0.00027778);

            decimal pos_ptht = (MLVSS / MLSS) * 100;
            decimal ol_mlss_aer = (sunoliko_v_dex_aer * MLSS) / 1000;
            decimal ol_mlvss_aer = (sunoliko_v_dex_aer * MLVSS) / 1000;
            decimal par_aera = BLOWERS_POWER * var * 24 * 50 * 24;
            decimal par_eis_viol = PAROXI_EISODOU - paroxi_a_il;
            decimal xronos_param_aer = sunoliko_v_dex_aer / (par_eis_viol / 24);
            decimal logos_aera = par_aera / par_eis_viol;
            decimal bod_apom = ((BOD_viol - BOD_out) * par_eis_viol) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal logosaer_pr_bodapom = par_aera / bod_apom;
            decimal apod_o2 = 0;
            if(logosaer_pr_bodapom != 0)
            {
                apod_o2 = logos_aera / logosaer_pr_bodapom;
            }
            decimal par_o2 = par_aera * apod_o2;
            decimal logoso2_pr_bodapom = par_o2 / bod_apom;
            decimal ogk_fort_bod = ((par_eis_viol * BOD_viol) / sunoliko_v_dex_aer) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal sun_pax_aer = par_eis_viol + (SUM - PAROXI_EISODOU);
            decimal vss_upol = bod_apom * VSS;
            decimal inerts_upol = ((100 - pos_ptht) * par_eis_viol * SS_viol) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal tss_upol = vss_upol + inerts_upol;
            decimal tss_diad = ((ol_mlss_aer + MLSS) * ogkos_bkath) * Convert.ToDecimal(Math.Pow(10, 3));
            decimal mcrt = tss_diad / tss_upol;
            decimal f_m = (par_eis_viol * BOD_viol) / ol_mlvss_aer;
            decimal n_nh3_eis = N_NH3_eis * par_eis_viol;
            decimal n_nh3_ex = N_NH3_ex * par_eis_viol;
            decimal n_no3_eis = N_NO3_eis * par_eis_viol;
            decimal n_no3_ex = N_NO3_ex * par_eis_viol;
            decimal n_nitro = -(n_nh3_ex - n_nh3_eis);
            decimal n_aponitro = (n_nitro + n_no3_eis) - n_no3_ex;
            decimal xron_aer = (ogkos_aerovias * 24) / par_eis_viol;
            decimal xron_aponitro = (ogkos_anox * 24) / par_eis_viol;
            decimal apod_n_nh3 = (N_NH3_eis - N_NH3_ex) / N_NH3_eis;
            decimal apod_n_no3 = (N_NO3_eis - N_NO3_ex) / N_NO3_eis;
            decimal xron_param = (ogkos_anaer * 24) / par_eis_viol;
            decimal apod_ptot = (Ptot_out - Ptot_out_dex_aer) / Ptot_out;

            command.CommandText = "insert into dex_aer(xronos_param_aer,pos_ptht,ol_mlss_aer,ol_mlvss_aer,par_aera,logos_aera,bod_apom,logosaer_pr_bodapom,apod_o2,par_o2,logoso2_pr_bodapom,ogk_fort_bod,sun_pax_aer,vss_upol,inerts_upol,tss_upol,tss_diad,mcrt,f_m,n_nh3_eis,n_no3_eis,n_nh3_ex,n_no3_ex,n_nitro,n_aponitro,xron_aer,xron_aponitro,apod_n_nh3,apod_n_no3,xron_param,apod_ptot,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13, @p14, @p15, @p16, @p17, @p18, @p19, @p20, @p21, @p22, @p23, @p24, @p25, @p26, @p27, @p28, @p29, @p30, @p31, @p32)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", xronos_param_aer);
            command.Parameters.AddWithValue("@p2", pos_ptht);
            command.Parameters.AddWithValue("@p3", ol_mlss_aer);
            command.Parameters.AddWithValue("@p4", ol_mlvss_aer);
            command.Parameters.AddWithValue("@p5", par_aera);
            command.Parameters.AddWithValue("@p6", logos_aera);
            command.Parameters.AddWithValue("@p7", bod_apom);
            command.Parameters.AddWithValue("@p8", logosaer_pr_bodapom);
            command.Parameters.AddWithValue("@p9", apod_o2);
            command.Parameters.AddWithValue("@p10", par_o2);
            command.Parameters.AddWithValue("@p11", logoso2_pr_bodapom);
            command.Parameters.AddWithValue("@p12", ogk_fort_bod);
            command.Parameters.AddWithValue("@p13", sun_pax_aer);
            command.Parameters.AddWithValue("@p14", vss_upol);
            command.Parameters.AddWithValue("@p15", inerts_upol);
            command.Parameters.AddWithValue("@p16", tss_upol);
            command.Parameters.AddWithValue("@p17", tss_diad);
            command.Parameters.AddWithValue("@p18", mcrt);
            command.Parameters.AddWithValue("@p19", f_m);
            command.Parameters.AddWithValue("@p20", n_nh3_eis);
            command.Parameters.AddWithValue("@p21", n_no3_eis);
            command.Parameters.AddWithValue("@p22", n_nh3_ex);
            command.Parameters.AddWithValue("@p23", n_no3_ex);
            command.Parameters.AddWithValue("@p24", n_nitro);
            command.Parameters.AddWithValue("@p25", n_aponitro);
            command.Parameters.AddWithValue("@p26", xron_aer);
            command.Parameters.AddWithValue("@p27", xron_aponitro);
            command.Parameters.AddWithValue("@p28", apod_n_nh3);
            command.Parameters.AddWithValue("@p29", apod_n_no3);
            command.Parameters.AddWithValue("@p30", xron_param);
            command.Parameters.AddWithValue("@p31", apod_ptot);
            command.Parameters.AddWithValue("@p32", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void b_kath(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal MLSS = Convert.ToDecimal(reader["MLSS"]);
            decimal MLVSS = Convert.ToDecimal(reader["MLVSS"]);
            decimal BOD_viol = Convert.ToDecimal(reader["BOD_viol"]);
            decimal SS_viol = Convert.ToDecimal(reader["SS_viol"]);
            decimal BOD_out = Convert.ToDecimal(reader["BOD_out"]);
            decimal VSS = Convert.ToDecimal(reader["VSS"]);
            decimal SS_out = Convert.ToDecimal(reader["SS_out"]);
            decimal N_NH3_ex = Convert.ToDecimal(reader["N_NH3_ex"]);
            decimal N_NO3_eis = Convert.ToDecimal(reader["N_NO3_eis"]);
            decimal N_NO3_ex = Convert.ToDecimal(reader["N_NO3_ex"]);
            decimal N_NH3_out = Convert.ToDecimal(reader["N_NH3_out"]);
            decimal N_NO3_out = Convert.ToDecimal(reader["N_NO3_out"]);
            decimal Ptot_out_kath = Convert.ToDecimal(reader["Ptot_out_kath"]);
            decimal Ptot_out_dex_aer = Convert.ToDecimal(reader["Ptot_out_dex_aer"]);
            decimal Ptot_in = Convert.ToDecimal(reader["Ptot_in"]);
            decimal Ptot_out = Convert.ToDecimal(reader["Ptot_out"]);
            
            reader.Close();

            command.CommandText = "SELECT * FROM dexamenes ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal sunoliko_v_dex_aer = Convert.ToDecimal(reader["sunoliko_v_dex_aer"]);
            decimal ogkos_bkath = Convert.ToDecimal(reader["ogkos_bkath"]);
            decimal ogkos_aerovias = Convert.ToDecimal(reader["ogkos_aerovias"]);
            decimal ogkos_anaer = Convert.ToDecimal(reader["ogkos_anaer"]);
            decimal epif_bkath = Convert.ToDecimal(reader["epif_bkath"]);

            reader.Close();

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal A3_FLOW = Convert.ToDecimal(reader["A3_FLOW"]);
            decimal PAROXI_EISODOU = Convert.ToDecimal(reader["PAROXI_EISODOU"]);
            decimal SUM = Convert.ToDecimal(reader["SUM"]);

            reader.Close();

            command.CommandText = "SELECT * FROM a_kath ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal paroxi_a_il = Convert.ToDecimal(reader["paroxi_a_il"]);

            reader.Close();

            command.CommandText = "SELECT * FROM dex_aer ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal n_nh3_ex = Convert.ToDecimal(reader["n_nh3_ex"]);
            decimal n_no3_ex = Convert.ToDecimal(reader["n_no3_ex"]);

            reader.Close();

            decimal par_eis_viol = PAROXI_EISODOU - paroxi_a_il;
            decimal sun_par = par_eis_viol + (SUM - PAROXI_EISODOU);
            decimal udrau_fort = par_eis_viol / epif_bkath;
            decimal fort_ster = ((sun_par * MLSS) / epif_bkath) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal xron_param = (ogkos_bkath * 24) / par_eis_viol;
            decimal fort_ss_out = (par_eis_viol * SS_out) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal bod_apom = ((BOD_viol - BOD_out) * par_eis_viol) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal bod_apod = (BOD_viol - BOD_out) / BOD_viol;
            decimal ss_apod = (SS_viol - SS_out) / SS_viol;
            decimal fort_n_nh3_out = (N_NH3_out * par_eis_viol) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal fort_n_no3_out = (N_NO3_out * par_eis_viol) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal n_nitro = n_nh3_ex - fort_n_nh3_out;
            decimal n_aponitro = (n_no3_ex - N_NO3_out) + n_nitro;
            decimal fort_ptot_out = (par_eis_viol * Ptot_out_kath) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal apod_ptot_b = (Ptot_out_dex_aer - Ptot_out_kath) / Ptot_out_dex_aer;
            decimal apod_ptot_a = (Ptot_in - Ptot_out) / Ptot_in;

            command.CommandText = "insert into b_kath(sun_par,udrau_fort,fort_ster,xron_param,fort_ss_out,bod_apom,bod_apod,ss_apod,fort_n_nh3_out,fort_n_no3_out,n_nitro,n_aponitro,fort_ptot_out,apod_ptot_b,apod_ptot_a,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13, @p14, @p15, @p16)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", sun_par);
            command.Parameters.AddWithValue("@p2", udrau_fort);
            command.Parameters.AddWithValue("@p3", fort_ster);
            command.Parameters.AddWithValue("@p4", xron_param);
            command.Parameters.AddWithValue("@p5", fort_ss_out);
            command.Parameters.AddWithValue("@p6", bod_apom);
            command.Parameters.AddWithValue("@p7", bod_apod);
            command.Parameters.AddWithValue("@p8", ss_apod);
            command.Parameters.AddWithValue("@p9", fort_n_nh3_out);
            command.Parameters.AddWithValue("@p10", fort_n_no3_out);
            command.Parameters.AddWithValue("@p11", n_nitro);
            command.Parameters.AddWithValue("@p12", n_aponitro);
            command.Parameters.AddWithValue("@p13", fort_ptot_out);
            command.Parameters.AddWithValue("@p14", apod_ptot_b);
            command.Parameters.AddWithValue("@p15", apod_ptot_a);
            command.Parameters.AddWithValue("@p16", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void c_vathm(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM a_kath ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal paroxi_a_il = Convert.ToDecimal(reader["paroxi_a_il"]);

            reader.Close();

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal PAROXI_EISODOU = Convert.ToDecimal(reader["PAROXI_EISODOU"]);

            reader.Close();

            command.CommandText = "SELECT * FROM dexamenes ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal epif_filtrwn = Convert.ToDecimal(reader["epif_filtrwn"]);

            reader.Close();

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal SS_viol = Convert.ToDecimal(reader["SS_viol"]);
            decimal SS_out = Convert.ToDecimal(reader["SS_out"]);
            decimal SS_ex = Convert.ToDecimal(reader["SS_ex"]);

            reader.Close();

            command.CommandText = "SELECT * FROM b_kath ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal fort_ss_out = Convert.ToDecimal(reader["fort_ss_out"]);

            reader.Close();

            decimal par_eis_viol = PAROXI_EISODOU - paroxi_a_il;
            decimal udrau_fort = par_eis_viol / epif_filtrwn;
            decimal fort_ster = (par_eis_viol * (SS_viol - fort_ss_out)) / epif_filtrwn;
            decimal apod_filtr = (SS_out - SS_ex) / SS_out;

            command.CommandText = "insert into c_vathm(udrau_fort,fort_ster,apod_filtr,date) values (@p1, @p2, @p3, @p4)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", udrau_fort);
            command.Parameters.AddWithValue("@p2", fort_ster);
            command.Parameters.AddWithValue("@p3", apod_filtr);
            command.Parameters.AddWithValue("@p4", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void ol_apod_eel(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal Ptot_in = Convert.ToDecimal(reader["Ptot_in"]);
            decimal Ptot_out_kath = Convert.ToDecimal(reader["Ptot_out_kath"]);
            decimal N_NH3_eis = Convert.ToDecimal(reader["N_NH3_eis"]);
            decimal N_NH3_out = Convert.ToDecimal(reader["N_NH3_out"]);
            decimal N_NO3_eis = Convert.ToDecimal(reader["N_NO3_eis"]);
            decimal N_NO3_out = Convert.ToDecimal(reader["N_NO3_out"]);
            decimal BOD = Convert.ToDecimal(reader["BOD"]);
            decimal BOD_ex = Convert.ToDecimal(reader["BOD_ex"]);
            decimal SS = Convert.ToDecimal(reader["SS"]);
            decimal SS_ex = Convert.ToDecimal(reader["SS_ex"]);
            decimal SS_viol = Convert.ToDecimal(reader["SS_viol"]);

            reader.Close();

            command.CommandText = "SELECT * FROM a_kath ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal fortio_ssa = Convert.ToDecimal(reader["fortio_ssa"]);

            reader.Close();

            command.CommandText = "SELECT * FROM c_vathm ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal apod_filtr = Convert.ToDecimal(reader["apod_filtr"]);

            reader.Close();

            decimal apod_ptot = (Ptot_in - Ptot_out_kath) / Ptot_in;
            decimal apod_n_nh3 = (N_NH3_eis - N_NH3_out) / N_NH3_eis;
            decimal apod_n_no3 = (N_NO3_eis - N_NO3_out) / N_NO3_eis;
            decimal apod_bod = (BOD - BOD_ex) / BOD;
            decimal apod_ss = (SS - SS_ex) / SS;
            decimal apod_ntot = 0;

            if (fortio_ssa != 0)
            {
                apod_ntot = (((fortio_ssa / SS_viol) * (N_NH3_eis - N_NH3_out)) + ((fortio_ssa / apod_filtr) * (N_NO3_eis - N_NO3_out))) / (((fortio_ssa / SS_viol) * N_NH3_eis) + ((fortio_ssa / apod_filtr) * N_NO3_eis)); 
            }

            command.CommandText = "insert into ol_apod_eel(apod_ptot,apod_n_nh3,apod_n_no3,apod_bod,apod_ss,apod_ntot,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", apod_ptot);
            command.Parameters.AddWithValue("@p2", apod_n_nh3);
            command.Parameters.AddWithValue("@p3", apod_n_no3);
            command.Parameters.AddWithValue("@p4", apod_bod);
            command.Parameters.AddWithValue("@p5", apod_ss);
            command.Parameters.AddWithValue("@p6", apod_ntot);
            command.Parameters.AddWithValue("@p7", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void propax(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal SS_propah = Convert.ToDecimal(reader["SS_propah"]);
            decimal SS_ex = Convert.ToDecimal(reader["SS_ex"]);
            
            reader.Close();

            command.CommandText = "SELECT * FROM dexamenes ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal ogkos_dex_propax = Convert.ToDecimal(reader["ogkos_dex_propax"]);
            decimal epif_dex_propax = Convert.ToDecimal(reader["epif_dex_propax"]);
            

            reader.Close();

            command.CommandText = "SELECT * FROM a_kath ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal paroxi_a_il = Convert.ToDecimal(reader["paroxi_a_il"]);
            decimal fortio_ssa = Convert.ToDecimal(reader["fortio_ssa"]);

            reader.Close();

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal PAROXH_A_ILYOS_1 = Convert.ToDecimal(reader["PAROXH_A_ILYOS_1"]);
            decimal PAROXH_A_ILYOS_3 = Convert.ToDecimal(reader["PAROXH_A_ILYOS_3"]);
            decimal PAROXH_PROP_1 = Convert.ToDecimal(reader["PAROXH_PROP_1"]);
            decimal PAROXH_PROP_2 = Convert.ToDecimal(reader["PAROXH_PROP_2"]);

            reader.Close();

            decimal xron_param = 0;
            if (paroxi_a_il != 0)
            {
                xron_param = (ogkos_dex_propax * 24) / paroxi_a_il; 
            }
            decimal udrau_fort = paroxi_a_il / epif_dex_propax;
            decimal fort_ster = ((SS_ex * paroxi_a_il) / epif_dex_propax) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal fort_ss_pax = ((PAROXH_PROP_1 + PAROXH_PROP_2) * SS_propah) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal fort_ss_stag = fortio_ssa - fort_ss_pax;
            decimal ss_stag = (fort_ss_stag / ((PAROXH_A_ILYOS_1 + PAROXH_A_ILYOS_3) - (PAROXH_PROP_1 + PAROXH_PROP_2))) * Convert.ToDecimal(Math.Pow(10, -3));

            command.CommandText = "insert into propax(xron_param,udrau_fort,fort_ster,fort_ss_pax,fort_ss_stag,ss_stag,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", xron_param);
            command.Parameters.AddWithValue("@p2", udrau_fort);
            command.Parameters.AddWithValue("@p3", fort_ster);
            command.Parameters.AddWithValue("@p4", fort_ss_pax);
            command.Parameters.AddWithValue("@p5", fort_ss_stag);
            command.Parameters.AddWithValue("@p6", ss_stag);
            command.Parameters.AddWithValue("@p7", date);
            command.ExecuteNonQuery();


            //connection.Close();
        }

        public static void mhx_pax(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal MLSS = Convert.ToDecimal(reader["MLSS"]);
            decimal MLSS_PERIS_NEW = Convert.ToDecimal(reader["MLSS_PERIS_NEW"]);
            decimal PERISIA_PAROXI = Convert.ToDecimal(reader["PERISIA_PAROXI"]);
            decimal FE02 = Convert.ToDecimal(reader["25IF02"]);
            decimal FLOW_OMOG = Convert.ToDecimal(reader["FLOW_OMOG"]);
            decimal PAROXH_PROP_1 = Convert.ToDecimal(reader["PAROXH_PROP_1"]);
            decimal PAROXH_PROP_2 = Convert.ToDecimal(reader["PAROXH_PROP_2"]);

            reader.Close();

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal SS_pah = Convert.ToDecimal(reader["SS_pah"]);
            decimal katan_PHL = Convert.ToDecimal(reader["katan_PHL"]);
            
            reader.Close();

            command.CommandText = "SELECT * FROM propax ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal fort_ss_propax = Convert.ToDecimal(reader["fort_ss_pax"]);
            
            reader.Close();
            
            decimal fort_wass = (((MLSS * 10000) + (MLSS_PERIS_NEW * 10000) / 2) * (PERISIA_PAROXI + FE02)) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal par_omog = FLOW_OMOG - (PAROXH_PROP_1 + PAROXH_PROP_2);
            decimal fort_ss_pax = (par_omog * SS_pah) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal fort_ss_strag = fort_wass - fort_ss_pax;
            decimal ss_strag = (fort_ss_strag / ((PERISIA_PAROXI + FE02) - (PAROXH_PROP_1 + PAROXH_PROP_2 + FLOW_OMOG))) * Convert.ToDecimal(Math.Pow(10, 3));
            decimal eid_kat_phl = katan_PHL / fort_ss_pax;
            decimal logos_pax = ((MLSS * 1000) + (MLSS_PERIS_NEW * 1000) / 2) / SS_pah;
            decimal fort_omog = fort_ss_pax + fort_ss_propax;
            decimal ss_pax = (fort_omog / FLOW_OMOG) * Convert.ToDecimal(Math.Pow(10, 3));

            command.CommandText = "insert into mhx_pax(fort_wass,par_omog,fort_ss_pax,fort_ss_strag,ss_strag,eid_kat_phl,logos_pax,fort_omog,ss_pax,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", fort_wass);
            command.Parameters.AddWithValue("@p2", par_omog);
            command.Parameters.AddWithValue("@p3", fort_ss_pax);
            command.Parameters.AddWithValue("@p4", fort_ss_strag);
            command.Parameters.AddWithValue("@p5", ss_strag);
            command.Parameters.AddWithValue("@p6", eid_kat_phl);
            command.Parameters.AddWithValue("@p7", logos_pax);
            command.Parameters.AddWithValue("@p8", fort_omog);
            command.Parameters.AddWithValue("@p9", ss_pax);
            command.Parameters.AddWithValue("@p10", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void xwneusi(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal FLOW_OMOG = Convert.ToDecimal(reader["FLOW_OMOG"]);

            reader.Close();

            command.CommandText = "SELECT * FROM dexamenes ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal ogkos_xwn = Convert.ToDecimal(reader["ogkos_xwn"]);

            reader.Close();

            command.CommandText = "SELECT * FROM mhx_pax ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal fort_omog = Convert.ToDecimal(reader["fort_omog"]);
            decimal ss_pax = Convert.ToDecimal(reader["ss_pax"]);

            reader.Close();

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal VSS_xon = Convert.ToDecimal(reader["VSS_xon"]);
            decimal SS_ex = Convert.ToDecimal(reader["SS_ex"]);
            decimal VSS_ex = Convert.ToDecimal(reader["VSS_ex"]);
            decimal VA = Convert.ToDecimal(reader["VA"]);
            decimal Alkalik = Convert.ToDecimal(reader["Alkalik"]);
            
            reader.Close();

            decimal xron_param = ogkos_xwn / FLOW_OMOG;
            decimal fort_ss_xwn = fort_omog;
            decimal fort_vss_xwn = (FLOW_OMOG * VSS_xon) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal pos_ptht_in = (VSS_xon / ss_pax) / 100;
            decimal fort_ptht = fort_vss_xwn / ogkos_xwn;
            decimal fort_ss_fix = fort_omog - fort_vss_xwn;
            decimal fort_ss_ex = (FLOW_OMOG * SS_ex) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal fort_vss_ex = (FLOW_OMOG * VSS_ex) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal pos_ptht_out = (VSS_ex / SS_ex) / 100;
            decimal parag_vioaer = fort_vss_xwn - fort_vss_ex;
            decimal apodosi = (fort_vss_xwn - fort_vss_ex) / fort_vss_xwn;
            decimal logos_va = VA / Alkalik;

            command.CommandText = "insert into xwneusi(xron_param,fort_ss_xwn,fort_vss_xwn,pos_ptht_in,fort_ptht,fort_ss_fix,fort_ss_ex,fort_vss_ex,pos_ptht_out,parag_vioaer,apodosi,logos_va,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", xron_param);
            command.Parameters.AddWithValue("@p2", fort_ss_xwn);
            command.Parameters.AddWithValue("@p3", fort_vss_xwn);
            command.Parameters.AddWithValue("@p4", pos_ptht_in);
            command.Parameters.AddWithValue("@p5", fort_ptht);
            command.Parameters.AddWithValue("@p6", fort_ss_fix);
            command.Parameters.AddWithValue("@p7", fort_ss_ex);
            command.Parameters.AddWithValue("@p8", fort_vss_ex);
            command.Parameters.AddWithValue("@p9", pos_ptht_out);
            command.Parameters.AddWithValue("@p10", parag_vioaer);
            command.Parameters.AddWithValue("@p11", apodosi);
            command.Parameters.AddWithValue("@p12", logos_va);
            command.Parameters.AddWithValue("@p13", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void metapax(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM dexamenes ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal epif_dex_metapax = Convert.ToDecimal(reader["epif_dex_metapax"]);
            decimal ogkos_dex_metapax = Convert.ToDecimal(reader["ogkos_dex_metapax"]);

            reader.Close();

            command.CommandText = "SELECT * FROM calculations ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal FLOW_OMOG = Convert.ToDecimal(reader["FLOW_OMOG"]);
            decimal FLOW = Convert.ToDecimal(reader["FLOW"]);
            decimal MLSS_DEC = Convert.ToDecimal(reader["MLSS_DEC"]);
            

            reader.Close();

            command.CommandText = "SELECT * FROM xwneusi ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal fort_ss_ex = Convert.ToDecimal(reader["fort_ss_ex"]);

            reader.Close();

            decimal xron_param = (ogkos_dex_metapax * 24) / FLOW_OMOG;
            decimal udrau_fort = FLOW_OMOG / epif_dex_metapax;
            decimal fort_ster = fort_ss_ex / epif_dex_metapax;
            decimal fort_ss_pax = ((MLSS_DEC * 10000) * FLOW) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal fort_ss_strag = fort_ss_ex - fort_ss_pax;
            decimal ss_strag = (fort_ss_strag / (FLOW_OMOG - FLOW)) * Convert.ToDecimal(Math.Pow(10, 3));

            command.CommandText = "insert into metapax(xron_param,udrau_fort,fort_ster,fort_ss_pax,fort_ss_strag,ss_strag,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", xron_param);
            command.Parameters.AddWithValue("@p2", udrau_fort);
            command.Parameters.AddWithValue("@p3", fort_ster);
            command.Parameters.AddWithValue("@p4", fort_ss_pax);
            command.Parameters.AddWithValue("@p5", fort_ss_strag);
            command.Parameters.AddWithValue("@p6", ss_strag);
            command.Parameters.AddWithValue("@p7", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }

        public static void afudat(DateTime date)
        {
            MySqlConnection connection = new MySqlConnection(sqlstringtext());
            connection.Open();
            MySqlCommand command = connection.CreateCommand();
            MySqlDataReader reader;

            command.CommandText = "SELECT * FROM inputs ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal SS_afud = Convert.ToDecimal(reader["SS_afud"]);
            
            reader.Close();

            command.CommandText = "SELECT * FROM metapax ORDER BY date DESC LIMIT 0,1";
            command.Prepare();
            reader = command.ExecuteReader();

            reader.Read();

            decimal fort_ss_pax = Convert.ToDecimal(reader["fort_ss_pax"]);

            reader.Close();

            decimal afudatwmenh = 1;
            decimal katan_phl = 1;

            decimal fort_ss_afud = (SS_afud * afudatwmenh) * Convert.ToDecimal(Math.Pow(10, -3));
            decimal fort_ss_strag = fort_ss_pax - fort_ss_afud;
            decimal eid_kat_phl = katan_phl / fort_ss_afud;

            command.CommandText = "insert into afudatwsi(fort_ss_afud,fort_ss_strag,eid_kat_phl,date) values (@p1, @p2, @p3, @p4)";
            command.Prepare();
            command.Parameters.AddWithValue("@p1", fort_ss_afud);
            command.Parameters.AddWithValue("@p2", fort_ss_strag);
            command.Parameters.AddWithValue("@p3", eid_kat_phl);
            command.Parameters.AddWithValue("@p4", date);
            command.ExecuteNonQuery();

            //connection.Close();
        }
    }
}
