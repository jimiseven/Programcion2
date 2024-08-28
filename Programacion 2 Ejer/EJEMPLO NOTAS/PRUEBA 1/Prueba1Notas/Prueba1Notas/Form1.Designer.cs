using System.ComponentModel;

namespace Prueba1NotasD
{
    partial class zlForm1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.ButonSeleccionarExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.ButtonGenerarWord = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ButonSeleccionarExcel
            // 
            this.ButonSeleccionarExcel.Location = new System.Drawing.Point(221, 178);
            this.ButonSeleccionarExcel.Name = "ButonSeleccionarExcel";
            this.ButonSeleccionarExcel.Size = new System.Drawing.Size(131, 23);
            this.ButonSeleccionarExcel.TabIndex = 0;
            this.ButonSeleccionarExcel.Text = "Seleccionar excel ";
            this.ButonSeleccionarExcel.UseVisualStyleBackColor = true;
            this.ButonSeleccionarExcel.Click += new System.EventHandler(this.ButonSeleccionarExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(184, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "GENERADOR DE BOLETINES ESCOLARES";
            // 
            // ButtonGenerarWord
            // 
            this.ButtonGenerarWord.Location = new System.Drawing.Point(246, 249);
            this.ButtonGenerarWord.Name = "ButtonGenerarWord";
            this.ButtonGenerarWord.Size = new System.Drawing.Size(75, 23);
            this.ButtonGenerarWord.TabIndex = 3;
            this.ButtonGenerarWord.Text = "GENERAR BOLETINES";
            this.ButtonGenerarWord.UseVisualStyleBackColor = true;
            this.ButtonGenerarWord.Click += new System.EventHandler(this.ButtonGenerarWord_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(615, 364);
            this.Controls.Add(this.ButtonGenerarWord);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ButonSeleccionarExcel);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ButonSeleccionarExcel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button ButtonGenerarWord;
    }
}

