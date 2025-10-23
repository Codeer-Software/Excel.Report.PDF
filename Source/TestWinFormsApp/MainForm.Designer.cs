namespace TestWinFormsApp
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            _buttonPreview = new Button();
            _buttonSettings = new Button();
            _textBoxFile = new TextBox();
            _buttonFile = new Button();
            SuspendLayout();
            // 
            // _buttonPreview
            // 
            _buttonPreview.Location = new Point(17, 48);
            _buttonPreview.Name = "_buttonPreview";
            _buttonPreview.Size = new Size(75, 23);
            _buttonPreview.TabIndex = 0;
            _buttonPreview.Text = "Prview";
            _buttonPreview.UseVisualStyleBackColor = true;
            _buttonPreview.Click += _buttonPreview_Click;
            // 
            // _buttonSettings
            // 
            _buttonSettings.Location = new Point(98, 48);
            _buttonSettings.Name = "_buttonSettings";
            _buttonSettings.Size = new Size(75, 23);
            _buttonSettings.TabIndex = 1;
            _buttonSettings.Text = "button2";
            _buttonSettings.UseVisualStyleBackColor = true;
            _buttonSettings.Click += _buttonSettings_Click;
            // 
            // _textBoxFile
            // 
            _textBoxFile.Location = new Point(17, 19);
            _textBoxFile.Name = "_textBoxFile";
            _textBoxFile.Size = new Size(323, 23);
            _textBoxFile.TabIndex = 2;
            // 
            // _buttonFile
            // 
            _buttonFile.Location = new Point(346, 19);
            _buttonFile.Name = "_buttonFile";
            _buttonFile.Size = new Size(29, 23);
            _buttonFile.TabIndex = 3;
            _buttonFile.Text = "...";
            _buttonFile.UseVisualStyleBackColor = true;
            _buttonFile.Click += _buttonFile_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(473, 98);
            Controls.Add(_buttonFile);
            Controls.Add(_textBoxFile);
            Controls.Add(_buttonSettings);
            Controls.Add(_buttonPreview);
            Name = "MainForm";
            Text = "Print";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button _buttonPreview;
        private Button _buttonSettings;
        private TextBox _textBoxFile;
        private Button _buttonFile;
    }
}
