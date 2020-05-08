using Microsoft.Toolkit.Forms.UI.Controls;

namespace Panel_Tracker
{
    partial class OAuth
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
            this.wv = new WebView();
            ((System.ComponentModel.ISupportInitialize)(this.wv)).BeginInit();
            this.SuspendLayout();
            // 
            // wv
            // 
            this.wv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wv.Location = new System.Drawing.Point(0, 0);
            this.wv.MinimumSize = new System.Drawing.Size(20, 20);
            this.wv.Name = "wv";
            this.wv.Size = new System.Drawing.Size(800, 450);
            this.wv.TabIndex = 0;
            // 
            // OAuth
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.wv);
            this.Name = "OAuth";
            this.Text = "OAuth";
            ((System.ComponentModel.ISupportInitialize)(this.wv)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Toolkit.Forms.UI.Controls.WebView wv;
    }
}