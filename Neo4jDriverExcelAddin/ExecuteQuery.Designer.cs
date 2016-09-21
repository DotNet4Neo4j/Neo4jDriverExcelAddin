namespace Neo4jDriverExcelAddin
{
    using System;

    partial class ExecuteQuery
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            if (disposing)
            {
                ExecuteCypher = null;
            }

            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtCypher = new System.Windows.Forms.TextBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.txtNeoResponse = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtCypher
            // 
            this.txtCypher.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCypher.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCypher.Location = new System.Drawing.Point(0, 0);
            this.txtCypher.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtCypher.Multiline = true;
            this.txtCypher.Name = "txtCypher";
            this.txtCypher.Size = new System.Drawing.Size(453, 233);
            this.txtCypher.TabIndex = 0;
            // 
            // btnExecute
            // 
            this.btnExecute.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExecute.Location = new System.Drawing.Point(323, 239);
            this.btnExecute.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(130, 48);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Execute Cypher";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // txtNeoResponse
            // 
            this.txtNeoResponse.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtNeoResponse.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNeoResponse.Location = new System.Drawing.Point(0, 292);
            this.txtNeoResponse.Multiline = true;
            this.txtNeoResponse.Name = "txtNeoResponse";
            this.txtNeoResponse.ReadOnly = true;
            this.txtNeoResponse.Size = new System.Drawing.Size(453, 206);
            this.txtNeoResponse.TabIndex = 2;
            // 
            // ExecuteQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.txtNeoResponse);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.txtCypher);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "ExecuteQuery";
            this.Size = new System.Drawing.Size(453, 498);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtCypher;
        private System.Windows.Forms.Button btnExecute;


        internal EventHandler<ExecuteCypherQueryArgs> ExecuteCypher;
        private System.Windows.Forms.TextBox txtNeoResponse;
    }
}
