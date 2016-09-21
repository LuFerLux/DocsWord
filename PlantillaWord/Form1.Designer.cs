namespace PlantillaWord
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
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
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtdni = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtempleado = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtdir = new System.Windows.Forms.TextBox();
            this.txtcargo = new System.Windows.Forms.TextBox();
            this.txtremuneracion = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtinicio = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtfin = new System.Windows.Forms.TextBox();
            this.btnimprimir = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtdni
            // 
            this.txtdni.Location = new System.Drawing.Point(117, 9);
            this.txtdni.MaxLength = 8;
            this.txtdni.Name = "txtdni";
            this.txtdni.Size = new System.Drawing.Size(111, 20);
            this.txtdni.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Apellidos y Nombres:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(73, 87);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Cargo:";
            // 
            // txtempleado
            // 
            this.txtempleado.Location = new System.Drawing.Point(117, 35);
            this.txtempleado.Name = "txtempleado";
            this.txtempleado.Size = new System.Drawing.Size(259, 20);
            this.txtempleado.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(43, 136);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Fecha Inicio:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(54, 165);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Fecha Fin:";
            // 
            // txtdir
            // 
            this.txtdir.Location = new System.Drawing.Point(117, 61);
            this.txtdir.Name = "txtdir";
            this.txtdir.Size = new System.Drawing.Size(259, 20);
            this.txtdir.TabIndex = 7;
            // 
            // txtcargo
            // 
            this.txtcargo.Location = new System.Drawing.Point(117, 84);
            this.txtcargo.Name = "txtcargo";
            this.txtcargo.Size = new System.Drawing.Size(157, 20);
            this.txtcargo.TabIndex = 8;
            // 
            // txtremuneracion
            // 
            this.txtremuneracion.Location = new System.Drawing.Point(117, 110);
            this.txtremuneracion.Name = "txtremuneracion";
            this.txtremuneracion.Size = new System.Drawing.Size(75, 20);
            this.txtremuneracion.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(82, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "DNI:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(59, 58);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Domicilio:";
            // 
            // txtinicio
            // 
            this.txtinicio.Location = new System.Drawing.Point(117, 136);
            this.txtinicio.Name = "txtinicio";
            this.txtinicio.Size = new System.Drawing.Size(100, 20);
            this.txtinicio.TabIndex = 12;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(32, 110);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(79, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "Remuneración:";
            // 
            // txtfin
            // 
            this.txtfin.Location = new System.Drawing.Point(117, 162);
            this.txtfin.Name = "txtfin";
            this.txtfin.Size = new System.Drawing.Size(100, 20);
            this.txtfin.TabIndex = 14;
            // 
            // btnimprimir
            // 
            this.btnimprimir.Location = new System.Drawing.Point(117, 199);
            this.btnimprimir.Name = "btnimprimir";
            this.btnimprimir.Size = new System.Drawing.Size(75, 23);
            this.btnimprimir.TabIndex = 15;
            this.btnimprimir.Text = "Imprimir";
            this.btnimprimir.UseVisualStyleBackColor = true;
            this.btnimprimir.Click += new System.EventHandler(this.btnimprimir_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(492, 80);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(127, 23);
            this.button1.TabIndex = 16;
            this.button1.Text = "Unificar Docs";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(651, 234);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnimprimir);
            this.Controls.Add(this.txtfin);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtinicio);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtremuneracion);
            this.Controls.Add(this.txtcargo);
            this.Controls.Add(this.txtdir);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtempleado);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtdni);
            this.Name = "Form1";
            this.Text = "Generar Contrato";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtdni;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtempleado;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtdir;
        private System.Windows.Forms.TextBox txtcargo;
        private System.Windows.Forms.TextBox txtremuneracion;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtinicio;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtfin;
        private System.Windows.Forms.Button btnimprimir;
        private System.Windows.Forms.Button button1;
    }
}

