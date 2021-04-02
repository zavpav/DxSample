namespace DyDocTestSS.Visual
{
    partial class DynamicSheet
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DynamicSheet));
            this.ssControl = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.ssControlFormulaBar = new DevExpress.XtraSpreadsheet.SpreadsheetFormulaBar();
            this.splitterControl1 = new DevExpress.XtraEditors.SplitterControl();
            this.ssControlToolTip = new DevExpress.Utils.ToolTipController(this.components);
            this.SuspendLayout();
            // 
            // ssControl
            // 
            this.ssControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ssControl.Location = new System.Drawing.Point(0, 29);
            this.ssControl.Name = "ssControl";
            this.ssControl.Options.Behavior.Column.Delete = DevExpress.XtraSpreadsheet.DocumentCapability.Hidden;
            this.ssControl.Options.Behavior.Column.Hide = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.ssControl.Options.Behavior.Column.Insert = DevExpress.XtraSpreadsheet.DocumentCapability.Hidden;
            this.ssControl.Options.Behavior.Column.Unhide = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.ssControl.Options.Behavior.MaxZoomFactor = 1000F;
            this.ssControl.Options.Behavior.MinZoomFactor = 0.6F;
            this.ssControl.Options.Behavior.Row.AutoFit = DevExpress.XtraSpreadsheet.DocumentCapability.Hidden;
            this.ssControl.Options.Behavior.Row.Delete = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.ssControl.Options.Behavior.Row.Hide = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.ssControl.Options.Behavior.Row.Insert = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.ssControl.Options.Behavior.Row.Unhide = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            this.ssControl.Options.Import.Csv.Encoding = ((System.Text.Encoding)(resources.GetObject("ssControl.Options.Import.Csv.Encoding")));
            this.ssControl.Options.Import.Txt.Encoding = ((System.Text.Encoding)(resources.GetObject("ssControl.Options.Import.Txt.Encoding")));
            this.ssControl.Size = new System.Drawing.Size(792, 526);
            this.ssControl.TabIndex = 1;
            this.ssControl.Text = "ssControl";
            this.ssControl.ToolTipController = this.ssControlToolTip;
            this.ssControl.PopupMenuShowing += new DevExpress.XtraSpreadsheet.PopupMenuShowingEventHandler(this.ssControl_PopupMenuShowing);
            this.ssControl.CustomDrawCell += new DevExpress.XtraSpreadsheet.CustomDrawCellEventHandler(this.ssControl_CustomDrawCell);
            this.ssControl.CustomDrawCellBackground += new DevExpress.XtraSpreadsheet.CustomDrawCellBackgroundEventHandler(this.ssControl_CustomDrawCellBackground);
            this.ssControl.CustomCellEdit += new DevExpress.XtraSpreadsheet.SpreadsheetCustomCellEditEventHandler(this.ssControl_CustomCellEdit);
            this.ssControl.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(this.ssControl_PropertyChanged);
            this.ssControl.DocumentLoaded += new System.EventHandler(this.ssControl_DocumentLoaded);
            this.ssControl.ContentChanged += new System.EventHandler(this.ssControl_ContentChanged);
            this.ssControl.CellBeginEdit += new DevExpress.XtraSpreadsheet.CellBeginEditEventHandler(this.ssControl_CellBeginEdit);
            this.ssControl.CellEndEdit += new DevExpress.XtraSpreadsheet.CellEndEditEventHandler(this.ssControl_CellEndEdit);
            this.ssControl.CellValueChanged += new DevExpress.XtraSpreadsheet.CellValueChangedEventHandler(this.ssControl_CellValueChanged);
            this.ssControl.RowsInserting += new DevExpress.Spreadsheet.RowsChangingEventHandler(this.ssControl_RowsInserting);
            this.ssControl.ColumnsRemoved += new DevExpress.Spreadsheet.ColumnsRemovedEventHandler(this.ssControl_ColumnsRemoved);
            this.ssControl.ColumnsRemoving += new DevExpress.Spreadsheet.ColumnsChangingEventHandler(this.ssControl_ColumnsRemoving);
            this.ssControl.ColumnsInserted += new DevExpress.Spreadsheet.ColumnsInsertedEventHandler(this.ssControl_ColumnsInserted);
            this.ssControl.ColumnsInserting += new DevExpress.Spreadsheet.ColumnsChangingEventHandler(this.ssControl_ColumnsInserting);
            this.ssControl.RangeCopying += new DevExpress.Spreadsheet.RangeCopyingEventHandler(this.ssControl_RangeCopying);
            this.ssControl.RangeCopied += new DevExpress.Spreadsheet.RangeCopiedEventHandler(this.ssControl_RangeCopied);
            this.ssControl.CopiedRangePasting += new DevExpress.Spreadsheet.CopiedRangePastingEventHandler(this.ssControl_CopiedRangePasting);
            this.ssControl.CopiedRangePasted += new DevExpress.Spreadsheet.CopiedRangePastedEventHandler(this.ssControl_CopiedRangePasted);
            this.ssControl.ClipboardDataPasting += new DevExpress.Spreadsheet.ClipboardDataPastingEventHandler(this.ssControl_ClipboardDataPasting);
            this.ssControl.ClipboardDataObtained += new DevExpress.Spreadsheet.ClipboardDataObtainedEventHandler(this.ssControl_ClipboardDataObtained);
            this.ssControl.ClipboardDataPasted += new DevExpress.Spreadsheet.ClipboardDataPastedEventHandler(this.ssControl_ClipboardDataPasted);
            // 
            // ssControlFormulaBar
            // 
            this.ssControlFormulaBar.Dock = System.Windows.Forms.DockStyle.Top;
            this.ssControlFormulaBar.Location = new System.Drawing.Point(0, 0);
            this.ssControlFormulaBar.MinimumSize = new System.Drawing.Size(0, 24);
            this.ssControlFormulaBar.Name = "ssControlFormulaBar";
            this.ssControlFormulaBar.Size = new System.Drawing.Size(792, 24);
            this.ssControlFormulaBar.SpreadsheetControl = this.ssControl;
            this.ssControlFormulaBar.TabIndex = 3;
            // 
            // splitterControl1
            // 
            this.splitterControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.splitterControl1.Location = new System.Drawing.Point(0, 24);
            this.splitterControl1.MinSize = 20;
            this.splitterControl1.Name = "splitterControl1";
            this.splitterControl1.Size = new System.Drawing.Size(792, 5);
            this.splitterControl1.TabIndex = 2;
            this.splitterControl1.TabStop = false;
            // 
            // ssControlToolTip
            // 
            this.ssControlToolTip.GetActiveObjectInfo += new DevExpress.Utils.ToolTipControllerGetActiveObjectInfoEventHandler(this.ssControlToolTip_GetActiveObjectInfo);
            // 
            // DynamicSheet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ssControl);
            this.Controls.Add(this.splitterControl1);
            this.Controls.Add(this.ssControlFormulaBar);
            this.Name = "DynamicSheet";
            this.Size = new System.Drawing.Size(792, 555);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraSpreadsheet.SpreadsheetControl ssControl;
        private DevExpress.XtraSpreadsheet.SpreadsheetFormulaBar ssControlFormulaBar;
        private DevExpress.XtraEditors.SplitterControl splitterControl1;
        private DevExpress.Utils.ToolTipController ssControlToolTip;
    }
}
