namespace 給料計算アプリサンプル
{
	partial class Att_04
	{
		/// <summary> 
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary> 
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region コンポーネント デザイナーで生成されたコード

		/// <summary> 
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.dataGridView1 = new System.Windows.Forms.DataGridView();
			this.button2 = new System.Windows.Forms.Button();
			this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			this.SuspendLayout();
			// 
			// dataGridView1
			// 
			this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView1.Location = new System.Drawing.Point(35, 25);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.RowTemplate.Height = 21;
			this.dataGridView1.Size = new System.Drawing.Size(733, 439);
			this.dataGridView1.TabIndex = 0;
			// 
			// button2
			// 
			this.button2.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.button2.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.button2.Location = new System.Drawing.Point(487, 548);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(282, 33);
			this.button2.TabIndex = 23;
			this.button2.Text = "Menyelesaikan informasi kehadiran.";
			this.button2.UseVisualStyleBackColor = false;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// dateTimePicker1
			// 
			this.dateTimePicker1.CalendarFont = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.dateTimePicker1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.dateTimePicker1.Location = new System.Drawing.Point(23, 555);
			this.dateTimePicker1.Name = "dateTimePicker1";
			this.dateTimePicker1.Size = new System.Drawing.Size(208, 23);
			this.dateTimePicker1.TabIndex = 24;
			// 
			// textBox1
			// 
			this.textBox1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.textBox1.Location = new System.Drawing.Point(316, 563);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(100, 23);
			this.textBox1.TabIndex = 25;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(272, 548);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(194, 12);
			this.label1.TabIndex = 26;
			this.label1.Text = "Jumlah makan siang yang diperlukan";
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.button1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.button1.Location = new System.Drawing.Point(174, 470);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(466, 33);
			this.button1.TabIndex = 27;
			this.button1.Text = "Mengajukan pertanyaan tentang ketidakhadiran yang tidak sah";
			this.button1.UseVisualStyleBackColor = false;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// Att_04
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.button1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.textBox1);
			this.Controls.Add(this.dateTimePicker1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.dataGridView1);
			this.Name = "Att_04";
			this.Size = new System.Drawing.Size(805, 630);
			this.Load += new System.EventHandler(this.Att_04_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.DateTimePicker dateTimePicker1;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
	}
}
