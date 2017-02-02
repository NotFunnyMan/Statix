using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Statix
{
    public partial class FormForErrors : MetroFramework.Forms.MetroForm
    {
        public FormForErrors()
        {
            InitializeComponent();
        }

        public FormForErrors(List<string> _errors)
        {
            InitializeComponent();

            for (int i = 0; i < _errors.Count; i++)
            {
                metroTextBox1.Text += (i + 1).ToString() + ".\t" + _errors[i] + Environment.NewLine;
                //ListViewItem lvi;
                //lvi = new ListViewItem();
                //lvi.SubItems.Add((i + 1).ToString());
                //lvi.SubItems.Add(_errors[i]);
            }
            
        }

        //Закрыть форму
        private void metroButton1_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
