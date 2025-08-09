namespace OlapTest
{
    partial class Form1
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
            
            if (disposing && (_dateFilterTimer != null))
            {
                _dateFilterTimer.Stop();
                _dateFilterTimer.Dispose();
                _dateFilterTimer = null;
            }
            
                            // Clear reference to active form when closing
            if (disposing)
            {
                                    // Use reflection to access static field
                try
                {
                    var field = typeof(Form1).GetField("_activeInstance", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                    if (field?.GetValue(null) == this)
                    {
                        field.SetValue(null, null);
                    }
                }
                catch
                {
                    // Ignore errors during cleanup
                }
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
            this.radPivotGrid1 = new Telerik.WinControls.UI.RadPivotGrid();
            this._statusStrip = new System.Windows.Forms.StatusStrip();
            this._statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.radPivotFieldList1 = new Telerik.WinControls.UI.RadPivotFieldList();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this._labelDateFrom = new System.Windows.Forms.Label();
            this._datePickerFrom = new Telerik.WinControls.UI.RadDateTimePicker();
            this._labelDateTo = new System.Windows.Forms.Label();
            this._datePickerTo = new Telerik.WinControls.UI.RadDateTimePicker();
            this._autoSaveLayout = new Telerik.WinControls.UI.RadCheckBox();
            this._updateButton = new Telerik.WinControls.UI.RadButton();
            this._exportXmlCheckBox = new Telerik.WinControls.UI.RadCheckBox();
            this._saveProfileButton = new Telerik.WinControls.UI.RadButton();
            this._loadProfileButton = new Telerik.WinControls.UI.RadButton();
            this._logTextBox = new System.Windows.Forms.RichTextBox();
            this._logSplitter = new System.Windows.Forms.Splitter();
            ((System.ComponentModel.ISupportInitialize)(this.radPivotGrid1)).BeginInit();
            this.radPivotGrid1.SuspendLayout();
            this._statusStrip.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this._datePickerFrom)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._datePickerTo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._autoSaveLayout)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._updateButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._exportXmlCheckBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._saveProfileButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._loadProfileButton)).BeginInit();
            this.SuspendLayout();
            // 
            // radPivotGrid1
            // 
            this.radPivotGrid1.ColumnWidth = 476;
            this.radPivotGrid1.Controls.Add(this._statusStrip);
            this.radPivotGrid1.DataMember = null;
            this.radPivotGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.radPivotGrid1.Location = new System.Drawing.Point(0, 42);
            this.radPivotGrid1.Margin = new System.Windows.Forms.Padding(12, 12, 12, 12);
            this.radPivotGrid1.Name = "radPivotGrid1";
            this.radPivotGrid1.ShowFilterArea = true;
            this.radPivotGrid1.Size = new System.Drawing.Size(1161, 438);
            this.radPivotGrid1.TabIndex = 0;
            // 
            // _statusStrip
            // 
            this._statusStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this._statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this._statusLabel});
            this._statusStrip.Location = new System.Drawing.Point(0, 412);
            this._statusStrip.Name = "_statusStrip";
            this._statusStrip.Padding = new System.Windows.Forms.Padding(1, 0, 28, 0);
            this._statusStrip.Size = new System.Drawing.Size(1161, 26);
            this._statusStrip.TabIndex = 0;
            this._statusStrip.Text = "_statusStrip";
            // 
            // _statusLabel
            // 
            this._statusLabel.Name = "_statusLabel";
            this._statusLabel.Size = new System.Drawing.Size(155, 20);
            this._statusLabel.Text = "Load time: 0 ms";
            // 
            // radPivotFieldList1
            // 
            this.radPivotFieldList1.AssociatedPivotGrid = this.radPivotGrid1;
            this.radPivotFieldList1.DeferUpdates = true;
            this.radPivotFieldList1.Dock = System.Windows.Forms.DockStyle.Right;
            this.radPivotFieldList1.Location = new System.Drawing.Point(1161, 0);
            this.radPivotFieldList1.MinimumSize = new System.Drawing.Size(225, 305);
            this.radPivotFieldList1.Name = "radPivotFieldList1";
            this.radPivotFieldList1.Size = new System.Drawing.Size(258, 480);
            this.radPivotFieldList1.TabIndex = 1;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.ColumnCount = 9;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this._updateButton, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this._labelDateFrom, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this._datePickerFrom, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this._labelDateTo, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this._datePickerTo, 4, 0);
            this.tableLayoutPanel1.Controls.Add(this._autoSaveLayout, 5, 0);
            this.tableLayoutPanel1.Controls.Add(this._exportXmlCheckBox, 6, 0);
            this.tableLayoutPanel1.Controls.Add(this._saveProfileButton, 7, 0);
            this.tableLayoutPanel1.Controls.Add(this._loadProfileButton, 8, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1161, 42);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // _labelDateFrom
            // 
            this._labelDateFrom.AutoSize = true;
            this._labelDateFrom.Dock = System.Windows.Forms.DockStyle.Fill;
            this._labelDateFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._labelDateFrom.Location = new System.Drawing.Point(3, 0);
            this._labelDateFrom.Name = "_labelDateFrom";
            this._labelDateFrom.Size = new System.Drawing.Size(93, 42);
            this._labelDateFrom.TabIndex = 1;
            this._labelDateFrom.Text = "Period from:";
            this._labelDateFrom.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // _datePickerFrom
            // 
            this._datePickerFrom.CustomFormat = "dd.MM.yyyy";
            this._datePickerFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this._datePickerFrom.Location = new System.Drawing.Point(104, 5);
            this._datePickerFrom.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this._datePickerFrom.Name = "_datePickerFrom";
            this._datePickerFrom.Size = new System.Drawing.Size(302, 32);
            this._datePickerFrom.TabIndex = 2;
            this._datePickerFrom.TabStop = false;
            this._datePickerFrom.Text = "01.01.2024";
            this._datePickerFrom.Value = new System.DateTime(2024, 1, 1, 0, 0, 0, 0);
            // 
            // _labelDateTo
            // 
            this._labelDateTo.AutoSize = true;
            this._labelDateTo.Dock = System.Windows.Forms.DockStyle.Fill;
            this._labelDateTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._labelDateTo.Location = new System.Drawing.Point(414, 0);
            this._labelDateTo.Name = "_labelDateTo";
            this._labelDateTo.Size = new System.Drawing.Size(34, 42);
            this._labelDateTo.TabIndex = 3;
            this._labelDateTo.Text = "to:";
            this._labelDateTo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // _datePickerTo
            // 
            this._datePickerTo.CustomFormat = "dd.MM.yyyy";
            this._datePickerTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this._datePickerTo.Location = new System.Drawing.Point(456, 5);
            this._datePickerTo.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this._datePickerTo.Name = "_datePickerTo";
            this._datePickerTo.Size = new System.Drawing.Size(302, 32);
            this._datePickerTo.TabIndex = 4;
            this._datePickerTo.TabStop = false;
            this._datePickerTo.Text = "31.12.2024";
            this._datePickerTo.Value = new System.DateTime(2024, 12, 31, 23, 59, 59, 999);
            // 
            // _updateButton
            // 
            this._updateButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._updateButton.Location = new System.Drawing.Point(768, 5);
            this._updateButton.Margin = new System.Windows.Forms.Padding(10, 5, 5, 5);
            this._updateButton.Name = "_updateButton";
            this._updateButton.Size = new System.Drawing.Size(100, 32);
            this._updateButton.TabIndex = 5;
            this._updateButton.Text = "Reload data";
            this._updateButton.Click += new System.EventHandler(this._updateButton_Click);
            // 
            // _autoSaveLayout
            // 
            this._autoSaveLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this._autoSaveLayout.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._autoSaveLayout.Location = new System.Drawing.Point(768, 10);
            this._autoSaveLayout.Margin = new System.Windows.Forms.Padding(5, 10, 5, 5);
            this._autoSaveLayout.Name = "_autoSaveLayout";
            this._autoSaveLayout.Size = new System.Drawing.Size(159, 23);
            this._autoSaveLayout.TabIndex = 0;
            this._autoSaveLayout.Text = "Auto save layout";
            this._autoSaveLayout.CheckStateChanged += new System.EventHandler(this._autoSaveLayout_CheckedChanged);
            // 
            // _exportXmlCheckBox
            // 
            this._exportXmlCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this._exportXmlCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._exportXmlCheckBox.Location = new System.Drawing.Point(933, 10);
            this._exportXmlCheckBox.Margin = new System.Windows.Forms.Padding(5, 10, 5, 5);
            this._exportXmlCheckBox.Name = "_exportXmlCheckBox";
            this._exportXmlCheckBox.Size = new System.Drawing.Size(159, 23);
            this._exportXmlCheckBox.TabIndex = 6;
            this._exportXmlCheckBox.Text = "Export XML profile";
            this._exportXmlCheckBox.CheckStateChanged += new System.EventHandler(this._exportXmlCheckBox_CheckedChanged);
            // 
            // _saveProfileButton
            // 
            this._saveProfileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._saveProfileButton.Location = new System.Drawing.Point(1102, 5);
            this._saveProfileButton.Margin = new System.Windows.Forms.Padding(10, 5, 5, 5);
            this._saveProfileButton.Name = "_saveProfileButton";
            this._saveProfileButton.Size = new System.Drawing.Size(100, 32);
            this._saveProfileButton.TabIndex = 7;
            this._saveProfileButton.Text = "Save profile";
            this._saveProfileButton.Click += new System.EventHandler(this._saveProfileButton_Click);
            // 
            // _loadProfileButton
            // 
            this._loadProfileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._loadProfileButton.Location = new System.Drawing.Point(1212, 5);
            this._loadProfileButton.Margin = new System.Windows.Forms.Padding(10, 5, 5, 5);
            this._loadProfileButton.Name = "_loadProfileButton";
            this._loadProfileButton.Size = new System.Drawing.Size(100, 32);
            this._loadProfileButton.TabIndex = 8;
            this._loadProfileButton.Text = "Load profile";
            this._loadProfileButton.Click += new System.EventHandler(this._loadProfileButton_Click);
            // 
            // _logTextBox
            // 
            this._logTextBox.BackColor = System.Drawing.Color.Black;
            this._logTextBox.Dock = System.Windows.Forms.DockStyle.Bottom;
            this._logTextBox.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._logTextBox.ForeColor = System.Drawing.Color.LimeGreen;
            this._logTextBox.Location = new System.Drawing.Point(0, 385);
            this._logTextBox.Name = "_logTextBox";
            this._logTextBox.ReadOnly = true;
            this._logTextBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this._logTextBox.Size = new System.Drawing.Size(1419, 100);
            this._logTextBox.TabIndex = 4;
            this._logTextBox.Text = "";
            // 
            // _logSplitter
            // 
            this._logSplitter.BackColor = System.Drawing.SystemColors.ControlDark;
            this._logSplitter.Cursor = System.Windows.Forms.Cursors.HSplit;
            this._logSplitter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this._logSplitter.Location = new System.Drawing.Point(0, 380);
            this._logSplitter.MinExtra = 100;
            this._logSplitter.MinSize = 50;
            this._logSplitter.Name = "_logSplitter";
            this._logSplitter.Size = new System.Drawing.Size(1419, 5);
            this._logSplitter.TabIndex = 3;
            this._logSplitter.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1419, 480);
            this.Controls.Add(this.radPivotGrid1);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this._logSplitter);
            this.Controls.Add(this._logTextBox);
            this.Controls.Add(this.radPivotFieldList1);
            this.Name = "Form1";
            this.Text = "Test tabular cube";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.radPivotGrid1)).EndInit();
            this.radPivotGrid1.ResumeLayout(false);
            this.radPivotGrid1.PerformLayout();
            this._statusStrip.ResumeLayout(false);
            this._statusStrip.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this._datePickerFrom)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._datePickerTo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._autoSaveLayout)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._updateButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._exportXmlCheckBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._saveProfileButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._loadProfileButton)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Telerik.WinControls.UI.RadPivotGrid radPivotGrid1;
        private Telerik.WinControls.UI.RadPivotFieldList radPivotFieldList1;
        private System.Windows.Forms.StatusStrip _statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel _statusLabel;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private Telerik.WinControls.UI.RadCheckBox _autoSaveLayout;
        private System.Windows.Forms.Label _labelDateFrom;
        private System.Windows.Forms.Label _labelDateTo;
        private Telerik.WinControls.UI.RadDateTimePicker _datePickerFrom;
        private Telerik.WinControls.UI.RadDateTimePicker _datePickerTo;
        private System.Windows.Forms.Timer _dateFilterTimer;
        private System.Windows.Forms.RichTextBox _logTextBox;
        private System.Windows.Forms.Splitter _logSplitter;
        private Telerik.WinControls.UI.RadButton _updateButton;
        private Telerik.WinControls.UI.RadCheckBox _exportXmlCheckBox;
        private Telerik.WinControls.UI.RadButton _saveProfileButton;
        private Telerik.WinControls.UI.RadButton _loadProfileButton;
    }
}