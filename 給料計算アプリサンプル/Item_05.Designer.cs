namespace 給料計算アプリサンプル
{
	partial class Item_05
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
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.dataGridView1 = new System.Windows.Forms.DataGridView();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.button13 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			this.SuspendLayout();
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("MS UI Gothic", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label1.Location = new System.Drawing.Point(74, 28);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 21);
			this.label1.TabIndex = 0;
			this.label1.Text = "<<";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("MS UI Gothic", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label2.Location = new System.Drawing.Point(165, 28);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(121, 21);
			this.label2.TabIndex = 1;
			this.label2.Text = "Maret 2025";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("MS UI Gothic", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label3.Location = new System.Drawing.Point(343, 28);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 21);
			this.label3.TabIndex = 2;
			this.label3.Text = ">>";
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// dataGridView1
			// 
			this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView1.Location = new System.Drawing.Point(470, 28);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.RowTemplate.Height = 21;
			this.dataGridView1.Size = new System.Drawing.Size(311, 513);
			this.dataGridView1.TabIndex = 3;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(36, 95);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 12);
			this.label4.TabIndex = 4;
			this.label4.Text = "Senin";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(88, 95);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 12);
			this.label5.TabIndex = 5;
			this.label5.Text = "Selasa";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(148, 95);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(31, 12);
			this.label6.TabIndex = 6;
			this.label6.Text = "Rabu";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(202, 95);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(36, 12);
			this.label7.TabIndex = 7;
			this.label7.Text = "Kamis";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(256, 95);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(37, 12);
			this.label8.TabIndex = 8;
			this.label8.Text = "Jumat";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(314, 95);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(33, 12);
			this.label9.TabIndex = 9;
			this.label9.Text = "sabtu";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(366, 95);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(41, 12);
			this.label10.TabIndex = 15;
			this.label10.Text = "minggu";
			// 
			// button13
			// 
			this.button13.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.button13.Font = new System.Drawing.Font("MS UI Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.button13.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
			this.button13.Location = new System.Drawing.Point(175, 483);
			this.button13.Name = "button13";
			this.button13.Size = new System.Drawing.Size(99, 37);
			this.button13.TabIndex = 23;
			this.button13.Text = "Daftar";
			this.button13.UseVisualStyleBackColor = false;
			this.button13.Click += new System.EventHandler(this.button13_Click);
			// 
			// Item_05
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.button13);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dataGridView1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Name = "Item_05";
			this.Size = new System.Drawing.Size(805, 630);
			this.Load += new System.EventHandler(this.Item_05_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.DataGridView dataGridView1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Button button13;
	}
}
