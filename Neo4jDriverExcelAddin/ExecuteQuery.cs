using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Neo4jDriverExcelAddin
{
    public partial class ExecuteQuery : UserControl
    {
        public ExecuteQuery()
        {
            InitializeComponent();
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            RaiseExecuteCypherEvent(txtCypher.Text);
        }

        private void RaiseExecuteCypherEvent(string cypher)
        {
            if (string.IsNullOrWhiteSpace(cypher))
                return;

            ExecuteCypher?.Invoke(this, new ExecuteCypherQueryArgs {Cypher = cypher});
        }
    }
}
