using System.Windows.Forms;
using System;
using System.Linq;

public partial class ChooseDateDialog : Form
{
    private DateTimePicker FromDatePicker;
    private DateTimePicker ToDatePicker;
    private Label FromLabel;
    private Label ToLabel;
    private Button CancelButton;
    private Button ContinueButton;

    public DateTime From { get; private set; }
    public DateTime To { get; private set; }

    public ChooseDateDialog(DateTime from, DateTime to)
    {
        StartPosition = FormStartPosition.Manual;
        Location = Cursor.Position;

        InitializeComponent();

        FromDatePicker.Value = from;
        ToDatePicker.Value = to;
    }

    private void ContinueButton_Click(object sender, EventArgs e)
    {
        From = FromDatePicker.Value;
        To = ToDatePicker.Value;
        this.DialogResult = DialogResult.OK;
        this.Close();
    }

    private void CancelButton_Click(object sender, EventArgs e)
    {
        this.DialogResult = DialogResult.Cancel;
        this.Close();
    }

    #region Not used

    private void InitializeComponent()
    {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChooseDateDialog));
            this.ToDatePicker = new System.Windows.Forms.DateTimePicker();
            this.FromDatePicker = new System.Windows.Forms.DateTimePicker();
            this.ToLabel = new System.Windows.Forms.Label();
            this.FromLabel = new System.Windows.Forms.Label();
            this.ContinueButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ToDatePicker
            // 
            resources.ApplyResources(this.ToDatePicker, "ToDatePicker");
            this.ToDatePicker.Name = "ToDatePicker";
            this.ToDatePicker.ValueChanged += new System.EventHandler(this.dateTimePicker2_ValueChanged);
            // 
            // FromDatePicker
            // 
            resources.ApplyResources(this.FromDatePicker, "FromDatePicker");
            this.FromDatePicker.Name = "FromDatePicker";
            // 
            // ToLabel
            // 
            resources.ApplyResources(this.ToLabel, "ToLabel");
            this.ToLabel.Name = "ToLabel";
            this.ToLabel.Click += new System.EventHandler(this.ToLabel_Click);
            // 
            // FromLabel
            // 
            resources.ApplyResources(this.FromLabel, "FromLabel");
            this.FromLabel.Name = "FromLabel";
            this.FromLabel.Click += new System.EventHandler(this.FromLabel_Click);
            // 
            // ContinueButton
            // 
            resources.ApplyResources(this.ContinueButton, "ContinueButton");
            this.ContinueButton.Name = "ContinueButton";
            this.ContinueButton.UseVisualStyleBackColor = true;
            this.ContinueButton.Click += new System.EventHandler(this.ContinueButton_Click);
            // 
            // CancelButton
            // 
            resources.ApplyResources(this.CancelButton, "CancelButton");
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ChooseDateDialog
            // 
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.ContinueButton);
            this.Controls.Add(this.ToLabel);
            this.Controls.Add(this.ToDatePicker);
            this.Controls.Add(this.FromDatePicker);
            this.Controls.Add(this.FromLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ChooseDateDialog";
            this.ShowIcon = false;
            this.Load += new System.EventHandler(this.InputDialog_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    private void InputDialog_Load(object sender, EventArgs e)
    {

    }

    private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
    {

    }

    private void label1_Click(object sender, EventArgs e)
    {

    }

    private void ToLabel_Click(object sender, EventArgs e)
    {

    }


    private void FromLabel_Click(object sender, EventArgs e)
    {

    }

    #endregion Not used
}