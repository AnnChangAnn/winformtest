
namespace MISSampleApp
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button_ImportExcel = new System.Windows.Forms.Button();
            this.button_GetExcel = new System.Windows.Forms.Button();
            this.textBox_FilePath = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button_InsertDB = new System.Windows.Forms.Button();
            this.button_SQLQuery = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button_ImportExcel
            // 
            this.button_ImportExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_ImportExcel.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_ImportExcel.Location = new System.Drawing.Point(15, 25);
            this.button_ImportExcel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button_ImportExcel.Name = "button_ImportExcel";
            this.button_ImportExcel.Size = new System.Drawing.Size(219, 39);
            this.button_ImportExcel.TabIndex = 0;
            this.button_ImportExcel.Text = "0.Open Excel";
            this.button_ImportExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_ImportExcel.UseVisualStyleBackColor = true;
            this.button_ImportExcel.Click += new System.EventHandler(this.buttonImport_Click);
            // 
            // button_GetExcel
            // 
            this.button_GetExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_GetExcel.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_GetExcel.Location = new System.Drawing.Point(15, 94);
            this.button_GetExcel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button_GetExcel.Name = "button_GetExcel";
            this.button_GetExcel.Size = new System.Drawing.Size(219, 39);
            this.button_GetExcel.TabIndex = 1;
            this.button_GetExcel.Text = "1.Import Excel";
            this.button_GetExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_GetExcel.UseVisualStyleBackColor = true;
            this.button_GetExcel.Click += new System.EventHandler(this.buttonGetData_Click);
            // 
            // textBox_FilePath
            // 
            this.textBox_FilePath.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.textBox_FilePath.Location = new System.Drawing.Point(261, 25);
            this.textBox_FilePath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox_FilePath.Name = "textBox_FilePath";
            this.textBox_FilePath.Size = new System.Drawing.Size(903, 34);
            this.textBox_FilePath.TabIndex = 4;
            this.textBox_FilePath.Text = "Excel File Path";
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(261, 94);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(903, 226);
            this.dataGridView1.TabIndex = 5;
            // 
            // button_InsertDB
            // 
            this.button_InsertDB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_InsertDB.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_InsertDB.Location = new System.Drawing.Point(15, 339);
            this.button_InsertDB.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button_InsertDB.Name = "button_InsertDB";
            this.button_InsertDB.Size = new System.Drawing.Size(219, 39);
            this.button_InsertDB.TabIndex = 7;
            this.button_InsertDB.Text = "2.Insert Data";
            this.button_InsertDB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_InsertDB.UseVisualStyleBackColor = true;
            this.button_InsertDB.Click += new System.EventHandler(this.button_InsertDB_Click);
            // 
            // button_SQLQuery
            // 
            this.button_SQLQuery.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_SQLQuery.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_SQLQuery.Location = new System.Drawing.Point(15, 404);
            this.button_SQLQuery.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button_SQLQuery.Name = "button_SQLQuery";
            this.button_SQLQuery.Size = new System.Drawing.Size(219, 39);
            this.button_SQLQuery.TabIndex = 8;
            this.button_SQLQuery.Text = "3. SQL Query";
            this.button_SQLQuery.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_SQLQuery.UseVisualStyleBackColor = true;
            this.button_SQLQuery.Click += new System.EventHandler(this.button_SQLQuery_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1176, 483);
            this.Controls.Add(this.button_SQLQuery);
            this.Controls.Add(this.button_InsertDB);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.textBox_FilePath);
            this.Controls.Add(this.button_GetExcel);
            this.Controls.Add(this.button_ImportExcel);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "MIS Sample APP";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_ImportExcel;
        private System.Windows.Forms.Button button_GetExcel;
        private System.Windows.Forms.TextBox textBox_FilePath;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button_InsertDB;
        private System.Windows.Forms.Button button_SQLQuery;
    }
}

