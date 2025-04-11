namespace 給料計算アプリサンプル
{
	partial class Cal_03
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
			this.button1 = new System.Windows.Forms.Button();
			this.label6 = new System.Windows.Forms.Label();
			this.dateTimePicker3 = new System.Windows.Forms.DateTimePicker();
			this.button2 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			this.SuspendLayout();
			// 
			// dataGridView1
			// 
			this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView1.Location = new System.Drawing.Point(24, 19);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.RowTemplate.Height = 21;
			this.dataGridView1.Size = new System.Drawing.Size(759, 468);
			this.dataGridView1.TabIndex = 0;
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.button1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.button1.Location = new System.Drawing.Point(484, 529);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(112, 31);
			this.button1.TabIndex = 19;
			this.button1.Text = "Impor data";
			this.button1.UseVisualStyleBackColor = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label6.Location = new System.Drawing.Point(165, 520);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(77, 16);
			this.label6.TabIndex = 27;
			this.label6.Text = "Hari gajian";
			// 
			// dateTimePicker3
			// 
			this.dateTimePicker3.CalendarFont = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.dateTimePicker3.Font = new System.Drawing.Font("MS UI Gothic", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.dateTimePicker3.Location = new System.Drawing.Point(118, 539);
			this.dateTimePicker3.Name = "dateTimePicker3";
			this.dateTimePicker3.Size = new System.Drawing.Size(178, 28);
			this.dateTimePicker3.TabIndex = 26;
			// 
			// button2
			// 
			this.button2.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.button2.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.button2.Location = new System.Drawing.Point(629, 529);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(112, 31);
			this.button2.TabIndex = 28;
			this.button2.Text = "Register";
			this.button2.UseVisualStyleBackColor = false;
			// 
			// Cal_03
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.button2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dateTimePicker3);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.dataGridView1);
			this.Name = "Cal_03";
			this.Size = new System.Drawing.Size(805, 630);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.DateTimePicker dateTimePicker3;
		private System.Windows.Forms.Button button2;
	}
}
