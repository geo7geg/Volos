using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Volos
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd-MM-yyyy";
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm:ss";
            dateTimePicker2.Value = DateTime.Now;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int seconds = Convert.ToInt32(dateTimePicker2.Value.ToString("ss"));
            int minutes = Convert.ToInt32(dateTimePicker2.Value.ToString("mm"));
            int hour = Convert.ToInt32(dateTimePicker2.Value.ToString("HH"));
            int dayofmonth = Convert.ToInt32(dateTimePicker1.Value.ToString("dd"));
            int month = Convert.ToInt32(dateTimePicker1.Value.ToString("MM"));
            int year = Convert.ToInt32(dateTimePicker1.Value.ToString("yyyy"));

            if (System.Windows.Forms.Application.OpenForms["Form1"] != null)
            {
                (System.Windows.Forms.Application.OpenForms["Form1"] as Form1).change_interval(seconds, minutes, hour, dayofmonth, month, year);
            }

            MessageBox.Show("Ο χρονοπρογραμματισμός ολοκληρώθηκε επιτυχώς για \n τις " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "\n και ώρα " + dateTimePicker2.Value.ToString("HH:mm:ss"));
        }
    }
}
