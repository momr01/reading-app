namespace MultasLectura.Views
{
    partial class GenerarTodo
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
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            txtCertificacion = new TextBox();
            btnCertificacion = new Button();
            txtPlazos = new TextBox();
            gbArchivos = new GroupBox();
            txtReclamos = new TextBox();
            btnReclamos = new Button();
            txtCalidadOperarios = new TextBox();
            btnCalidadOperarios = new Button();
            txtCalidad = new TextBox();
            btnCalidad = new Button();
            btnPlazos = new Button();
            btnGenerarLibros = new Button();
            gbBaremos = new GroupBox();
            lblEstadoBaremos = new Label();
            btnCargarBaremos = new Button();
            baremosAlturaT3 = new Label();
            baremosAlturaT1 = new Label();
            baremosT3 = new Label();
            label8 = new Label();
            label7 = new Label();
            label6 = new Label();
            baremosT2 = new Label();
            label4 = new Label();
            baremosT1 = new Label();
            label2 = new Label();
            gbMetas = new GroupBox();
            lblMetasAPI = new Label();
            btnCargarMetas = new Button();
            meta2 = new Label();
            meta1 = new Label();
            label10 = new Label();
            label9 = new Label();
            gbFechas = new GroupBox();
            lblFechasRes = new Label();
            label3 = new Label();
            label1 = new Label();
            dateTimePicker2 = new DateTimePicker();
            dateTimePicker1 = new DateTimePicker();
            lblFechasSeleccionadas = new Label();
            gbArchivos.SuspendLayout();
            gbBaremos.SuspendLayout();
            gbMetas.SuspendLayout();
            gbFechas.SuspendLayout();
            SuspendLayout();
            // 
            // txtCertificacion
            // 
            txtCertificacion.AllowDrop = true;
            txtCertificacion.Location = new Point(6, 43);
            txtCertificacion.Name = "txtCertificacion";
            txtCertificacion.ReadOnly = true;
            txtCertificacion.Size = new Size(364, 29);
            txtCertificacion.TabIndex = 0;
            txtCertificacion.TextChanged += txtCertificacion_TextChanged;
            txtCertificacion.DragDrop += txtCertificacion_DragDrop;
            txtCertificacion.DragEnter += txtCertificacion_DragEnter;
            // 
            // btnCertificacion
            // 
            btnCertificacion.BackColor = SystemColors.ActiveCaption;
            btnCertificacion.Location = new Point(376, 25);
            btnCertificacion.Name = "btnCertificacion";
            btnCertificacion.Size = new Size(196, 63);
            btnCertificacion.TabIndex = 1;
            btnCertificacion.Text = "Certificación detalle";
            btnCertificacion.UseVisualStyleBackColor = false;
            btnCertificacion.Click += btnCertificacion_Click;
            // 
            // txtPlazos
            // 
            txtPlazos.Location = new Point(6, 123);
            txtPlazos.Name = "txtPlazos";
            txtPlazos.ReadOnly = true;
            txtPlazos.Size = new Size(364, 29);
            txtPlazos.TabIndex = 2;
            txtPlazos.TextChanged += txtPlazos_TextChanged;
            txtPlazos.DragDrop += txtPlazos_DragDrop;
            txtPlazos.DragEnter += txtPlazos_DragEnter;
            // 
            // gbArchivos
            // 
            gbArchivos.Controls.Add(txtReclamos);
            gbArchivos.Controls.Add(btnReclamos);
            gbArchivos.Controls.Add(txtCalidadOperarios);
            gbArchivos.Controls.Add(btnCalidadOperarios);
            gbArchivos.Controls.Add(txtCalidad);
            gbArchivos.Controls.Add(btnCalidad);
            gbArchivos.Controls.Add(btnPlazos);
            gbArchivos.Controls.Add(btnCertificacion);
            gbArchivos.Controls.Add(txtCertificacion);
            gbArchivos.Controls.Add(txtPlazos);
            gbArchivos.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            gbArchivos.Location = new Point(12, 12);
            gbArchivos.Name = "gbArchivos";
            gbArchivos.Size = new Size(578, 451);
            gbArchivos.TabIndex = 4;
            gbArchivos.TabStop = false;
            gbArchivos.Text = "Archivos";
            // 
            // txtReclamos
            // 
            txtReclamos.Location = new Point(6, 387);
            txtReclamos.Name = "txtReclamos";
            txtReclamos.ReadOnly = true;
            txtReclamos.Size = new Size(364, 29);
            txtReclamos.TabIndex = 9;
            txtReclamos.TextChanged += txtReclamos_TextChanged;
            txtReclamos.DragDrop += txtReclamos_DragDrop;
            txtReclamos.DragEnter += txtReclamos_DragEnter;
            // 
            // btnReclamos
            // 
            btnReclamos.BackColor = SystemColors.ActiveCaption;
            btnReclamos.Location = new Point(376, 369);
            btnReclamos.Name = "btnReclamos";
            btnReclamos.Size = new Size(196, 63);
            btnReclamos.TabIndex = 8;
            btnReclamos.Text = "Reclamos detalle";
            btnReclamos.UseVisualStyleBackColor = false;
            btnReclamos.Click += btnReclamos_Click;
            // 
            // txtCalidadOperarios
            // 
            txtCalidadOperarios.Location = new Point(6, 299);
            txtCalidadOperarios.Name = "txtCalidadOperarios";
            txtCalidadOperarios.ReadOnly = true;
            txtCalidadOperarios.Size = new Size(364, 29);
            txtCalidadOperarios.TabIndex = 7;
            txtCalidadOperarios.TextChanged += txtCalidadOperarios_TextChanged;
            txtCalidadOperarios.DragDrop += txtCalidadOperarios_DragDrop;
            txtCalidadOperarios.DragEnter += txtCalidadOperarios_DragEnter;
            // 
            // btnCalidadOperarios
            // 
            btnCalidadOperarios.BackColor = SystemColors.ActiveCaption;
            btnCalidadOperarios.Location = new Point(376, 281);
            btnCalidadOperarios.Name = "btnCalidadOperarios";
            btnCalidadOperarios.Size = new Size(196, 63);
            btnCalidadOperarios.TabIndex = 6;
            btnCalidadOperarios.Text = "Calidad Operarios detalle";
            btnCalidadOperarios.UseVisualStyleBackColor = false;
            btnCalidadOperarios.Click += btnCalidadOperarios_Click;
            // 
            // txtCalidad
            // 
            txtCalidad.Location = new Point(6, 211);
            txtCalidad.Name = "txtCalidad";
            txtCalidad.ReadOnly = true;
            txtCalidad.Size = new Size(364, 29);
            txtCalidad.TabIndex = 5;
            txtCalidad.TextChanged += txtCalidad_TextChanged;
            txtCalidad.DragDrop += txtCalidad_DragDrop;
            txtCalidad.DragEnter += txtCalidad_DragEnter;
            // 
            // btnCalidad
            // 
            btnCalidad.BackColor = SystemColors.ActiveCaption;
            btnCalidad.Location = new Point(376, 193);
            btnCalidad.Name = "btnCalidad";
            btnCalidad.Size = new Size(196, 63);
            btnCalidad.TabIndex = 4;
            btnCalidad.Text = "Calidad detalle";
            btnCalidad.UseVisualStyleBackColor = false;
            btnCalidad.Click += btnCalidad_Click;
            // 
            // btnPlazos
            // 
            btnPlazos.BackColor = SystemColors.ActiveCaption;
            btnPlazos.Location = new Point(376, 105);
            btnPlazos.Name = "btnPlazos";
            btnPlazos.Size = new Size(196, 63);
            btnPlazos.TabIndex = 3;
            btnPlazos.Text = "Plazos detalle";
            btnPlazos.UseVisualStyleBackColor = false;
            btnPlazos.Click += btnPlazos_Click;
            // 
            // btnGenerarLibros
            // 
            btnGenerarLibros.BackColor = SystemColors.HotTrack;
            btnGenerarLibros.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btnGenerarLibros.ForeColor = Color.White;
            btnGenerarLibros.Location = new Point(12, 617);
            btnGenerarLibros.Name = "btnGenerarLibros";
            btnGenerarLibros.Size = new Size(853, 50);
            btnGenerarLibros.TabIndex = 5;
            btnGenerarLibros.Text = "GENERAR ARCHIVOS";
            btnGenerarLibros.UseVisualStyleBackColor = false;
            btnGenerarLibros.Click += btnGenerarLibros_Click;
            // 
            // gbBaremos
            // 
            gbBaremos.Controls.Add(lblEstadoBaremos);
            gbBaremos.Controls.Add(btnCargarBaremos);
            gbBaremos.Controls.Add(baremosAlturaT3);
            gbBaremos.Controls.Add(baremosAlturaT1);
            gbBaremos.Controls.Add(baremosT3);
            gbBaremos.Controls.Add(label8);
            gbBaremos.Controls.Add(label7);
            gbBaremos.Controls.Add(label6);
            gbBaremos.Controls.Add(baremosT2);
            gbBaremos.Controls.Add(label4);
            gbBaremos.Controls.Add(baremosT1);
            gbBaremos.Controls.Add(label2);
            gbBaremos.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            gbBaremos.Location = new Point(596, 12);
            gbBaremos.Name = "gbBaremos";
            gbBaremos.Size = new Size(269, 256);
            gbBaremos.TabIndex = 6;
            gbBaremos.TabStop = false;
            gbBaremos.Text = "Baremos";
            // 
            // lblEstadoBaremos
            // 
            lblEstadoBaremos.AutoSize = true;
            lblEstadoBaremos.Location = new Point(6, 219);
            lblEstadoBaremos.Name = "lblEstadoBaremos";
            lblEstadoBaremos.Size = new Size(0, 21);
            lblEstadoBaremos.TabIndex = 10;
            // 
            // btnCargarBaremos
            // 
            btnCargarBaremos.Location = new Point(6, 171);
            btnCargarBaremos.Name = "btnCargarBaremos";
            btnCargarBaremos.Size = new Size(257, 45);
            btnCargarBaremos.TabIndex = 9;
            btnCargarBaremos.Text = "Cargar";
            btnCargarBaremos.UseVisualStyleBackColor = true;
            // 
            // baremosAlturaT3
            // 
            baremosAlturaT3.AutoSize = true;
            baremosAlturaT3.Location = new Point(109, 134);
            baremosAlturaT3.Name = "baremosAlturaT3";
            baremosAlturaT3.Size = new Size(19, 21);
            baremosAlturaT3.TabIndex = 8;
            baremosAlturaT3.Text = "0";
            // 
            // baremosAlturaT1
            // 
            baremosAlturaT1.AutoSize = true;
            baremosAlturaT1.Location = new Point(110, 105);
            baremosAlturaT1.Name = "baremosAlturaT1";
            baremosAlturaT1.Size = new Size(19, 21);
            baremosAlturaT1.TabIndex = 8;
            baremosAlturaT1.Text = "0";
            // 
            // baremosT3
            // 
            baremosT3.AutoSize = true;
            baremosT3.Location = new Point(50, 79);
            baremosT3.Name = "baremosT3";
            baremosT3.Size = new Size(19, 21);
            baremosT3.TabIndex = 7;
            baremosT3.Text = "0";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(5, 134);
            label8.Name = "label8";
            label8.Size = new Size(98, 21);
            label8.TabIndex = 6;
            label8.Text = "ALTURA T3=";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(6, 106);
            label7.Name = "label7";
            label7.Size = new Size(98, 21);
            label7.TabIndex = 5;
            label7.Text = "ALTURA T1=";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(6, 79);
            label6.Name = "label6";
            label6.Size = new Size(38, 21);
            label6.TabIndex = 4;
            label6.Text = "T3=";
            // 
            // baremosT2
            // 
            baremosT2.AutoSize = true;
            baremosT2.Location = new Point(50, 52);
            baremosT2.Name = "baremosT2";
            baremosT2.Size = new Size(19, 21);
            baremosT2.TabIndex = 3;
            baremosT2.Text = "0";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(6, 52);
            label4.Name = "label4";
            label4.Size = new Size(38, 21);
            label4.TabIndex = 2;
            label4.Text = "T2=";
            // 
            // baremosT1
            // 
            baremosT1.AutoSize = true;
            baremosT1.Location = new Point(50, 25);
            baremosT1.Name = "baremosT1";
            baremosT1.Size = new Size(19, 21);
            baremosT1.TabIndex = 1;
            baremosT1.Text = "0";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(6, 25);
            label2.Name = "label2";
            label2.Size = new Size(38, 21);
            label2.TabIndex = 0;
            label2.Text = "T1=";
            // 
            // gbMetas
            // 
            gbMetas.Controls.Add(lblMetasAPI);
            gbMetas.Controls.Add(btnCargarMetas);
            gbMetas.Controls.Add(meta2);
            gbMetas.Controls.Add(meta1);
            gbMetas.Controls.Add(label10);
            gbMetas.Controls.Add(label9);
            gbMetas.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            gbMetas.Location = new Point(596, 274);
            gbMetas.Name = "gbMetas";
            gbMetas.Size = new Size(269, 189);
            gbMetas.TabIndex = 7;
            gbMetas.TabStop = false;
            gbMetas.Text = "Metas";
            // 
            // lblMetasAPI
            // 
            lblMetasAPI.AutoSize = true;
            lblMetasAPI.Location = new Point(6, 149);
            lblMetasAPI.Name = "lblMetasAPI";
            lblMetasAPI.Size = new Size(0, 21);
            lblMetasAPI.TabIndex = 12;
            // 
            // btnCargarMetas
            // 
            btnCargarMetas.Location = new Point(6, 101);
            btnCargarMetas.Name = "btnCargarMetas";
            btnCargarMetas.Size = new Size(257, 45);
            btnCargarMetas.TabIndex = 11;
            btnCargarMetas.Text = "Cargar";
            btnCargarMetas.UseVisualStyleBackColor = true;
            // 
            // meta2
            // 
            meta2.AutoSize = true;
            meta2.Location = new Point(85, 61);
            meta2.Name = "meta2";
            meta2.Size = new Size(36, 21);
            meta2.TabIndex = 9;
            meta2.Text = "0 %";
            // 
            // meta1
            // 
            meta1.AutoSize = true;
            meta1.Location = new Point(85, 31);
            meta1.Name = "meta1";
            meta1.Size = new Size(36, 21);
            meta1.TabIndex = 8;
            meta1.Text = "0 %";
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Location = new Point(6, 62);
            label10.Name = "label10";
            label10.Size = new Size(73, 21);
            label10.TabIndex = 1;
            label10.Text = "META 2=";
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(6, 31);
            label9.Name = "label9";
            label9.Size = new Size(73, 21);
            label9.TabIndex = 0;
            label9.Text = "META 1=";
            // 
            // gbFechas
            // 
            gbFechas.Controls.Add(lblFechasRes);
            gbFechas.Controls.Add(label3);
            gbFechas.Controls.Add(label1);
            gbFechas.Controls.Add(dateTimePicker2);
            gbFechas.Controls.Add(dateTimePicker1);
            gbFechas.Controls.Add(lblFechasSeleccionadas);
            gbFechas.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            gbFechas.Location = new Point(127, 469);
            gbFechas.Name = "gbFechas";
            gbFechas.Size = new Size(590, 142);
            gbFechas.TabIndex = 9;
            gbFechas.TabStop = false;
            gbFechas.Text = "Fecha Certificación";
            // 
            // lblFechasRes
            // 
            lblFechasRes.AutoSize = true;
            lblFechasRes.Location = new Point(125, 83);
            lblFechasRes.Name = "lblFechasRes";
            lblFechasRes.Size = new Size(0, 21);
            lblFechasRes.TabIndex = 16;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(318, 36);
            label3.Name = "label3";
            label3.Size = new Size(52, 21);
            label3.TabIndex = 15;
            label3.Text = "Hasta:";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(6, 36);
            label1.Name = "label1";
            label1.Size = new Size(56, 21);
            label1.TabIndex = 14;
            label1.Text = "Desde:";
            // 
            // dateTimePicker2
            // 
            dateTimePicker2.Location = new Point(376, 30);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(200, 29);
            dateTimePicker2.TabIndex = 13;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(68, 30);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(200, 29);
            dateTimePicker1.TabIndex = 12;
            // 
            // lblFechasSeleccionadas
            // 
            lblFechasSeleccionadas.AutoSize = true;
            lblFechasSeleccionadas.Location = new Point(6, 83);
            lblFechasSeleccionadas.Name = "lblFechasSeleccionadas";
            lblFechasSeleccionadas.Size = new Size(113, 21);
            lblFechasSeleccionadas.TabIndex = 11;
            lblFechasSeleccionadas.Text = "Rango elegido:";
            // 
            // GenerarTodo
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(877, 679);
            Controls.Add(gbFechas);
            Controls.Add(gbArchivos);
            Controls.Add(gbMetas);
            Controls.Add(gbBaremos);
            Controls.Add(btnGenerarLibros);
            Name = "GenerarTodo";
            Text = "Generar Todo";
            gbArchivos.ResumeLayout(false);
            gbArchivos.PerformLayout();
            gbBaremos.ResumeLayout(false);
            gbBaremos.PerformLayout();
            gbMetas.ResumeLayout(false);
            gbMetas.PerformLayout();
            gbFechas.ResumeLayout(false);
            gbFechas.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TextBox txtCertificacion;
        private Button btnCertificacion;
        private TextBox txtPlazos;
        private GroupBox gbArchivos;
        private TextBox txtCalidad;
        private Button btnCalidad;
        private Button btnPlazos;
        private Button btnGenerarLibros;
        private GroupBox gbBaremos;
        private Label label8;
        private Label label7;
        private Label label6;
        private Label baremosT2;
        private Label label4;
        private Label baremosT1;
        private Label label2;
        private GroupBox gbMetas;
        private Label label10;
        private Label label9;
        private Label baremosAlturaT3;
        private Label baremosAlturaT1;
        private Label baremosT3;
        private Label meta2;
        private Label meta1;
        private TextBox txtCalidadOperarios;
        private Button btnCalidadOperarios;
        private TextBox txtReclamos;
        private Button btnReclamos;
        private MonthCalendar monthCalendar1;
        private GroupBox gbFechas;
        private MonthCalendar monthCalendar2;
        private Label lblEstadoBaremos;
        private Button btnCargarBaremos;
        private Label lblMetasAPI;
        private Button btnCargarMetas;
        private Label lblFechasSeleccionadas;
        private Label lblFechasRes;
        private Label label3;
        private Label label1;
        private DateTimePicker dateTimePicker2;
        private DateTimePicker dateTimePicker1;
    }
}