using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAutosave
{
    public class TaskPaneControl : UserControl
    {
        private CheckBox toggleCheckBox;
        private Label statusLabel;
        private Label timeLabel;
        private Label infoLabel;
        private Timer autoSaveTimer;

        public TaskPaneControl()
        {
            InitializeComponent();
            SetupTimer();
        }

        private void InitializeComponent()
        {
            this.toggleCheckBox = new CheckBox();
            this.statusLabel = new Label();
            this.timeLabel = new Label();
            this.infoLabel = new Label();
            this.SuspendLayout();

            // 
            // toggleCheckBox
            // 
            this.toggleCheckBox.AutoSize = true;
            this.toggleCheckBox.Location = new Point(15, 20);
            this.toggleCheckBox.Text = "Enable Auto-Save";
            this.toggleCheckBox.Checked = true;
            this.toggleCheckBox.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            this.toggleCheckBox.CheckedChanged += new EventHandler(this.toggleCheckBox_CheckedChanged);

            // 
            // statusLabel
            // 
            this.statusLabel.Location = new Point(15, 60);
            this.statusLabel.Size = new Size(250, 20);
            this.statusLabel.Text = "Auto-saving is active...";
            this.statusLabel.ForeColor = Color.Green;
            this.statusLabel.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));

            // 
            // timeLabel
            // 
            this.timeLabel.Location = new Point(15, 85);
            this.timeLabel.Size = new Size(250, 20);
            this.timeLabel.Text = "Last saved: Not yet";
            this.timeLabel.ForeColor = Color.DimGray;
            this.timeLabel.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

            // 
            // infoLabel
            // 
            this.infoLabel.Location = new Point(15, 120);
            this.infoLabel.Size = new Size(250, 60);
            this.infoLabel.Text = "Checks for unsaved changes every 10 seconds and saves silently in the background.";
            this.infoLabel.ForeColor = Color.DimGray;
            this.infoLabel.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

            // 
            // TaskPaneControl
            // 
            this.Controls.Add(this.toggleCheckBox);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.timeLabel);
            this.Controls.Add(this.infoLabel);
            this.Size = new Size(300, 200);
            this.BackColor = Color.White;
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void SetupTimer()
        {
            autoSaveTimer = new Timer();
            autoSaveTimer.Interval = 10000; // 10 seconds
            autoSaveTimer.Tick += AutoSaveTimer_Tick;
            autoSaveTimer.Start();
        }

        private void toggleCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (toggleCheckBox.Checked)
            {
                statusLabel.Text = "Auto-saving is active...";
                statusLabel.ForeColor = Color.Green;
                autoSaveTimer.Start();
            }
            else
            {
                statusLabel.Text = "Auto-saving paused.";
                statusLabel.ForeColor = Color.Red;
                autoSaveTimer.Stop();
            }
        }

        private void AutoSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                if (app != null && app.Documents.Count > 0)
                {
                    Word.Document doc = app.ActiveDocument;
                    if (!doc.Saved)
                    {
                        // Check if it has a file path. If it's a completely new, unsaved document, Path is empty.
                        if (!string.IsNullOrEmpty(doc.Path))
                        {
                            doc.Save();
                            timeLabel.Text = "Last saved at " + DateTime.Now.ToString("HH:mm:ss");
                        }
                    }
                }
            }
            catch (Exception)
            {
                statusLabel.Text = "Error saving document.";
                statusLabel.ForeColor = Color.Red;
            }
        }
    }
}
