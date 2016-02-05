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
using MySql.Data.MySqlClient;
using MySql.Data;

namespace Volos
{
    public partial class Form3 : Form
    {
        public Form3()
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

                command.CommandText = "insert into dexamenes (epif_akath,ogkos_akath,ogkos_dex_aer_v,sunoliko_v_dex_aer,ogkos_bkath,ogkos_anox,ogkos_aerovias,ogkos_anaer,epif_bkath,mhkos_uper,epif_filtrwn,epif_dex_propax,ogkos_dex_propax,ogkos_xwn,epif_dex_metapax,ogkos_dex_metapax,date) values (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13, @p14, @p15, @p16, @p17)";
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
                command.Parameters.AddWithValue("@p17", dateTimePicker1.Value);
                command.ExecuteNonQuery();

                MessageBox.Show("Η εκχώρηση των τιμών ολοκληρώθηκε");

                connection.Close();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Σφάλμα επικοινωνίας με διακομιστή", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }
    }
}
