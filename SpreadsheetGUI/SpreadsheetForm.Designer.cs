namespace SpreadsheetGUI
{
    partial class SpreadsheetForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SpreadsheetForm));
            this.spreadsheetPanel1 = new SS.SpreadsheetPanel();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewHelpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip2 = new System.Windows.Forms.ToolStrip();
            this.errorText = new SpreadsheetGUI.ToolStripLabelSpring();
            this.cellNameTextBox = new System.Windows.Forms.ToolStripTextBox();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.cellContentsTextBox = new SpreadsheetGUI.ToolStripTextboxSpring();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.cellValueTextBox = new System.Windows.Forms.ToolStripTextBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.twitterButton = new System.Windows.Forms.ToolStripSplitButton();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.twitterStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.menuStrip1.SuspendLayout();
            this.toolStrip2.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // spreadsheetPanel1
            // 
            this.spreadsheetPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetPanel1.Location = new System.Drawing.Point(0, 74);
            this.spreadsheetPanel1.Name = "spreadsheetPanel1";
            this.spreadsheetPanel1.Size = new System.Drawing.Size(422, 269);
            this.spreadsheetPanel1.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(422, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem,
            this.openToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            this.closeToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // newToolStripMenuItem
            // 
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.Size = new System.Drawing.Size(123, 22);
            this.newToolStripMenuItem.Text = "New";
            this.newToolStripMenuItem.Click += new System.EventHandler(this.newToolStripMenuItem_Click);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(123, 22);
            this.openToolStripMenuItem.Text = "Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(123, 22);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(123, 22);
            this.saveAsToolStripMenuItem.Text = "Save As...";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.saveAsToolStripMenuItem_Click);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(123, 22);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.viewHelpToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // viewHelpToolStripMenuItem
            // 
            this.viewHelpToolStripMenuItem.Name = "viewHelpToolStripMenuItem";
            this.viewHelpToolStripMenuItem.Size = new System.Drawing.Size(127, 22);
            this.viewHelpToolStripMenuItem.Text = "View Help";
            this.viewHelpToolStripMenuItem.Click += new System.EventHandler(this.viewHelpToolStripMenuItem_Click);
            // 
            // toolStrip2
            // 
            this.toolStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.errorText});
            this.toolStrip2.Location = new System.Drawing.Point(0, 49);
            this.toolStrip2.Name = "toolStrip2";
            this.toolStrip2.Size = new System.Drawing.Size(422, 25);
            this.toolStrip2.Stretch = true;
            this.toolStrip2.TabIndex = 3;
            this.toolStrip2.Text = "toolStrip2";
            // 
            // errorText
            // 
            this.errorText.ForeColor = System.Drawing.Color.Red;
            this.errorText.Name = "errorText";
            this.errorText.Size = new System.Drawing.Size(381, 22);
            this.errorText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cellNameTextBox
            // 
            this.cellNameTextBox.Name = "cellNameTextBox";
            this.cellNameTextBox.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never;
            this.cellNameTextBox.ReadOnly = true;
            this.cellNameTextBox.Size = new System.Drawing.Size(50, 25);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never;
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cellNameTextBox,
            this.toolStripSeparator1,
            this.cellContentsTextBox,
            this.toolStripSeparator2,
            this.cellValueTextBox});
            this.toolStrip1.Location = new System.Drawing.Point(0, 24);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(422, 25);
            this.toolStrip1.Stretch = true;
            this.toolStrip1.TabIndex = 2;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // cellContentsTextBox
            // 
            this.cellContentsTextBox.Name = "cellContentsTextBox";
            this.cellContentsTextBox.Size = new System.Drawing.Size(213, 25);
            this.cellContentsTextBox.Leave += new System.EventHandler(this.cellContentsTextBox_Leave);
            this.cellContentsTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cellContentsTextBox_KeyDown);
            this.cellContentsTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cellContentsTextBox_KeyPress);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // cellValueTextBox
            // 
            this.cellValueTextBox.Name = "cellValueTextBox";
            this.cellValueTextBox.ReadOnly = true;
            this.cellValueTextBox.Size = new System.Drawing.Size(100, 25);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.twitterButton,
            this.toolStripStatusLabel1,
            this.twitterStatusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 343);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(422, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // twitterButton
            // 
            this.twitterButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.twitterButton.Image = ((System.Drawing.Image)(resources.GetObject("twitterButton.Image")));
            this.twitterButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.twitterButton.Name = "twitterButton";
            this.twitterButton.Size = new System.Drawing.Size(159, 20);
            this.twitterButton.Text = "Enable Twitter Integration";
            this.twitterButton.ButtonClick += new System.EventHandler(this.twitterButton_ButtonClick);
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // twitterStatusLabel
            // 
            this.twitterStatusLabel.Name = "twitterStatusLabel";
            this.twitterStatusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // SpreadsheetForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 365);
            this.Controls.Add(this.spreadsheetPanel1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.toolStrip2);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "SpreadsheetForm";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip2.ResumeLayout(false);
            this.toolStrip2.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private SS.SpreadsheetPanel spreadsheetPanel1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewHelpToolStripMenuItem;
        private System.Windows.Forms.ToolStrip toolStrip2;
        private ToolStripLabelSpring errorText;
        private System.Windows.Forms.ToolStripTextBox cellNameTextBox;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private ToolStripTextboxSpring cellContentsTextBox;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripTextBox cellValueTextBox;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripSplitButton twitterButton;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel twitterStatusLabel;
    }
}

