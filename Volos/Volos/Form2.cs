using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using MySql.Data;

namespace Volos
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd-MM-yyyy";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlConnection connection = new MySqlConnection(ClassLibrary.Class1.sqlstringtext());
                connection.Open();
                MySqlCommand command = connection.CreateCommand();

                command.CommandText = "insert into inputs(BOD,SS,Ptot_in,Ptot_out,BOD_viol,SS_viol,MLSS,MLVSS,VSS,N_NH3_eis,N_NO3_eis,N_NH3_ex,N_NO3_ex,Ptot_out_dex_aer,SS_out,BOD_out,N_NH3_out,N_NO3_out,Ptot_out_kath,BOD_ex,SS_ex,SS_propah,SS_pah,katan_PHL,Paroxi_xon,VSS_xon,VSS_ex,Alkalik,pH,VA,CH4,CO2,SS_afud,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13, @p14, @p15, @p16, @p17, @p18, @p19, @p20, @p21, @p22, @p23, @p24, @p25, @p26, @p27, @p28, @p29, @p30, @p31, @p32, @p33, @p34)";
                command.Prepare();
                command.Parameters.AddWithValue("@p1", Convert.ToDecimal(textBox1.Text));
                command.Parameters.AddWithValue("@p2", Convert.ToDecimal(textBox2.Text));
                command.Parameters.AddWithValue("@p3", Convert.ToDecimal(textBox3.Text));
                command.Parameters.AddWithValue("@p4", Convert.ToDecimal(textBox4.Text));
                command.Parameters.AddWithValue("@p5", Convert.ToDecimal(textBox5.Text));
                command.Parameters.AddWithValue("@p6", Convert.ToDecimal(textBox6.Text));
                command.Parameters.AddWithValue("@p7", Convert.ToDecimal(textBox7.Text));
                command.Parameters.AddWithValue("@p8", Convert.ToDecimal(textBox8.Text));
                command.Parameters.AddWithValue("@p9", Convert.ToDecimal(textBox9.Text));
                command.Parameters.AddWithValue("@p10", Convert.ToDecimal(textBox10.Text));
                command.Parameters.AddWithValue("@p11", Convert.ToDecimal(textBox11.Text));
                command.Parameters.AddWithValue("@p12", Convert.ToDecimal(textBox12.Text));
                command.Parameters.AddWithValue("@p13", Convert.ToDecimal(textBox13.Text));
                command.Parameters.AddWithValue("@p14", Convert.ToDecimal(textBox14.Text));
                command.Parameters.AddWithValue("@p15", Convert.ToDecimal(textBox15.Text));
                command.Parameters.AddWithValue("@p16", Convert.ToDecimal(textBox16.Text));
                command.Parameters.AddWithValue("@p17", Convert.ToDecimal(textBox17.Text));
                command.Parameters.AddWithValue("@p18", Convert.ToDecimal(textBox18.Text));
                command.Parameters.AddWithValue("@p19", Convert.ToDecimal(textBox19.Text));
                command.Parameters.AddWithValue("@p20", Convert.ToDecimal(textBox20.Text));
                command.Parameters.AddWithValue("@p21", Convert.ToDecimal(textBox21.Text));
                command.Parameters.AddWithValue("@p22", Convert.ToDecimal(textBox22.Text));
                command.Parameters.AddWithValue("@p23", Convert.ToDecimal(textBox23.Text));
                command.Parameters.AddWithValue("@p24", Convert.ToDecimal(textBox24.Text));
                command.Parameters.AddWithValue("@p25", Convert.ToDecimal(textBox25.Text));
                command.Parameters.AddWithValue("@p26", Convert.ToDecimal(textBox26.Text));
                command.Parameters.AddWithValue("@p27", Convert.ToDecimal(textBox27.Text));
                command.Parameters.AddWithValue("@p28", Convert.ToDecimal(textBox28.Text));
                command.Parameters.AddWithValue("@p29", Convert.ToDecimal(textBox29.Text));
                command.Parameters.AddWithValue("@p30", Convert.ToDecimal(textBox30.Text));
                command.Parameters.AddWithValue("@p31", Convert.ToDecimal(textBox31.Text));
                command.Parameters.AddWithValue("@p32", Convert.ToDecimal(textBox32.Text));
                command.Parameters.AddWithValue("@p33", Convert.ToDecimal(textBox33.Text));
                command.Parameters.AddWithValue("@p34", dateTimePicker1.Value);
                command.ExecuteNonQuery();

                MessageBox.Show("Η εκχώρηση των τιμών ολοκληρώθηκε");

                connection.Close();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Σφάλμα επικοινωνίας με διακομιστή", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
