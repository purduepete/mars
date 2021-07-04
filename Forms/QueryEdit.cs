using System;
using System.Windows.Forms;

namespace MarsAddIn.Forms
{
    public partial class QueryEdit : Form
    {
        public QueryEdit(string query = "")
        {
            InitializeComponent();
            txtQuery.Text = query;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var prnt = (DatabaseEntry)this.Owner;
            prnt.CustomQuery = txtQuery.Text;
            prnt.QueryEditorManipulations(txtQuery.Text);

            Hide();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            var prnt = (DatabaseEntry)this.Owner;

            txtQuery.Text = "";
            prnt.QueryEditorManipulations();

            Hide();
        }
    }
}
