namespace InquireCardTool
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "123",
            "456",
            "789"}, -1);
            this.ComPort_Label = new System.Windows.Forms.Label();
            this.ComPort_ComboBox = new System.Windows.Forms.ComboBox();
            this.Content_ListView = new System.Windows.Forms.ListView();
            this.CardItem = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Value = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.LastValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.RS232 = new System.IO.Ports.SerialPort(this.components);
            this.Info_RichTextBox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // ComPort_Label
            // 
            this.ComPort_Label.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ComPort_Label.AutoSize = true;
            this.ComPort_Label.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ComPort_Label.Location = new System.Drawing.Point(22, 24);
            this.ComPort_Label.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.ComPort_Label.Name = "ComPort_Label";
            this.ComPort_Label.Size = new System.Drawing.Size(75, 16);
            this.ComPort_Label.TabIndex = 0;
            this.ComPort_Label.Text = "ComPort : ";
            // 
            // ComPort_ComboBox
            // 
            this.ComPort_ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComPort_ComboBox.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ComPort_ComboBox.FormattingEnabled = true;
            this.ComPort_ComboBox.Items.AddRange(new object[] {
            "選擇ComPort"});
            this.ComPort_ComboBox.Location = new System.Drawing.Point(90, 19);
            this.ComPort_ComboBox.Margin = new System.Windows.Forms.Padding(2);
            this.ComPort_ComboBox.Name = "ComPort_ComboBox";
            this.ComPort_ComboBox.Size = new System.Drawing.Size(114, 24);
            this.ComPort_ComboBox.TabIndex = 1;
            this.ComPort_ComboBox.Tag = "";
            this.ComPort_ComboBox.DropDown += new System.EventHandler(this.ComPort_ComboBox_DropDown);
            this.ComPort_ComboBox.TextChanged += new System.EventHandler(this.ComPort_ComboBox_TextChanged);
            // 
            // Content_ListView
            // 
            this.Content_ListView.BackColor = System.Drawing.Color.White;
            this.Content_ListView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Content_ListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.CardItem,
            this.Value,
            this.LastValue});
            this.Content_ListView.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Content_ListView.FullRowSelect = true;
            this.Content_ListView.GridLines = true;
            this.Content_ListView.HideSelection = false;
            listViewItem1.IndentCount = 1;
            listViewItem1.StateImageIndex = 0;
            this.Content_ListView.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1});
            this.Content_ListView.Location = new System.Drawing.Point(22, 64);
            this.Content_ListView.Margin = new System.Windows.Forms.Padding(2);
            this.Content_ListView.Name = "Content_ListView";
            this.Content_ListView.Size = new System.Drawing.Size(601, 481);
            this.Content_ListView.TabIndex = 2;
            this.Content_ListView.UseCompatibleStateImageBehavior = false;
            this.Content_ListView.View = System.Windows.Forms.View.Details;
            // 
            // CardItem
            // 
            this.CardItem.Text = "項目";
            this.CardItem.Width = 200;
            // 
            // Value
            // 
            this.Value.Text = "本次";
            this.Value.Width = 180;
            // 
            // LastValue
            // 
            this.LastValue.Text = "前次";
            this.LastValue.Width = 180;
            // 
            // RS232
            // 
            this.RS232.BaudRate = 57600;
            // 
            // Info_RichTextBox
            // 
            this.Info_RichTextBox.Location = new System.Drawing.Point(630, 64);
            this.Info_RichTextBox.Margin = new System.Windows.Forms.Padding(2);
            this.Info_RichTextBox.Name = "Info_RichTextBox";
            this.Info_RichTextBox.ReadOnly = true;
            this.Info_RichTextBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth;
            this.Info_RichTextBox.Size = new System.Drawing.Size(188, 481);
            this.Info_RichTextBox.TabIndex = 3;
            this.Info_RichTextBox.Text = "";
            this.Info_RichTextBox.TextChanged += new System.EventHandler(this.Info_RichTextBox_TextChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(842, 562);
            this.Controls.Add(this.Info_RichTextBox);
            this.Controls.Add(this.Content_ListView);
            this.Controls.Add(this.ComPort_ComboBox);
            this.Controls.Add(this.ComPort_Label);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "MainForm";
            this.Text = "InquireCardTool";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label ComPort_Label;
        private System.Windows.Forms.ListView Content_ListView;
        private System.Windows.Forms.ColumnHeader CardItem;
        private System.Windows.Forms.ColumnHeader Value;
        private System.Windows.Forms.ColumnHeader LastValue;
        public System.Windows.Forms.ComboBox ComPort_ComboBox;
        private System.IO.Ports.SerialPort RS232;
        private System.Windows.Forms.RichTextBox Info_RichTextBox;
    }
}

