namespace Neo4jDriverExcelAddin
{
    using System;
    using System.Windows.Forms;

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

        public void SetMessage(string message)
        {
            txtNeoResponse.Text = message;
        }
    }
}