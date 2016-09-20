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
            this.SuspendLayout();
            // 
            // txtCypher
            // 
            this.txtCypher.Location = new System.Drawing.Point(0, 0);
            this.txtCypher.Multiline = true;
            this.txtCypher.Name = "txtCypher";
            this.txtCypher.Size = new System.Drawing.Size(237, 154);
            this.txtCypher.TabIndex = 0;
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(162, 160);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(75, 23);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Execute";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // ExecuteQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.txtCypher);
            this.Name = "ExecuteQuery";
            this.Size = new System.Drawing.Size(237, 244);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtCypher;
        private System.Windows.Forms.Button btnExecute;


        internal EventHandler<ExecuteCypherQueryArgs> ExecuteCypher;
    }
}
