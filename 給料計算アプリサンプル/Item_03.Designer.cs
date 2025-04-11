namespace 給料計算アプリサンプル
{
	partial class Item_03
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
			this.label1 = new System.Windows.Forms.Label();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.textBox2 = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.textBox3 = new System.Windows.Forms.TextBox();
			this.button1 = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.textBox4 = new System.Windows.Forms.TextBox();
			this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label1.Location = new System.Drawing.Point(45, 101);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(90, 16);
			this.label1.TabIndex = 13;
			this.label1.Text = "BPJS KESH";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// textBox1
			// 
			this.textBox1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.textBox1.Location = new System.Drawing.Point(172, 98);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(137, 23);
			this.textBox1.TabIndex = 12;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label2.Location = new System.Drawing.Point(45, 136);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(100, 16);
			this.label2.TabIndex = 15;
			this.label2.Text = "BPJS NAKER";
			// 
			// textBox2
			// 
			this.textBox2.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.textBox2.Location = new System.Drawing.Point(172, 133);
			this.textBox2.Name = "textBox2";
			this.textBox2.Size = new System.Drawing.Size(137, 23);
			this.textBox2.TabIndex = 14;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label3.Location = new System.Drawing.Point(45, 168);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(114, 16);
			this.label3.TabIndex = 17;
			this.label3.Text = "BPJS PENSIUN";
			// 
			// textBox3
			// 
			this.textBox3.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.textBox3.Location = new System.Drawing.Point(172, 165);
			this.textBox3.Name = "textBox3";
			this.textBox3.Size = new System.Drawing.Size(137, 23);
			this.textBox3.TabIndex = 16;
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.button1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.button1.Location = new System.Drawing.Point(634, 179);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(78, 31);
			this.button1.TabIndex = 19;
			this.button1.Text = "Daftar";
			this.button1.UseVisualStyleBackColor = false;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label4.Location = new System.Drawing.Point(347, 186);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(185, 16);
			this.label4.TabIndex = 31;
			this.label4.Text = "Tanggal mulai yang berlaku";
			// 
			// textBox4
			// 
			this.textBox4.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.textBox4.Location = new System.Drawing.Point(350, 205);
			this.textBox4.Name = "textBox4";
			this.textBox4.Size = new System.Drawing.Size(199, 23);
			this.textBox4.TabIndex = 30;
			// 
			// monthCalendar1
			// 
			this.monthCalendar1.Location = new System.Drawing.Point(350, 9);
			this.monthCalendar1.Name = "monthCalendar1";
			this.monthCalendar1.TabIndex = 29;
			this.monthCalendar1.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateChanged);
			// 
			// Item_03
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.label4);
			this.Controls.Add(this.textBox4);
			this.Controls.Add(this.monthCalendar1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.textBox3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.textBox2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.textBox1);
			this.Name = "Item_03";
			this.Size = new System.Drawing.Size(805, 630);
			this.Load += new System.EventHandler(this.Item_03_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox textBox2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textBox3;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBox4;
		private System.Windows.Forms.MonthCalendar monthCalendar1;
	}
}
