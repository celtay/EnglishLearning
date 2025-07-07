namespace EngWordSelect
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
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnNext = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.txtWord = new System.Windows.Forms.TextBox();
            this.txtTranslation = new System.Windows.Forms.TextBox();
            this.txtSentence1 = new System.Windows.Forms.TextBox();
            this.txtSentence2 = new System.Windows.Forms.TextBox();
            this.txtSentence3 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.shwButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(30, 112);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(120, 50);
            this.btnNext.TabIndex = 0;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(30, 177);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(120, 50);
            this.btnReset.TabIndex = 1;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // txtWord
            // 
            this.txtWord.Location = new System.Drawing.Point(228, 60);
            this.txtWord.Name = "txtWord";
            this.txtWord.Size = new System.Drawing.Size(236, 20);
            this.txtWord.TabIndex = 2;
            // 
            // txtTranslation
            // 
            this.txtTranslation.Location = new System.Drawing.Point(487, 60);
            this.txtTranslation.Name = "txtTranslation";
            this.txtTranslation.Size = new System.Drawing.Size(236, 20);
            this.txtTranslation.TabIndex = 3;
            // 
            // txtSentence1
            // 
            this.txtSentence1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSentence1.Location = new System.Drawing.Point(354, 124);
            this.txtSentence1.Name = "txtSentence1";
            this.txtSentence1.Size = new System.Drawing.Size(418, 20);
            this.txtSentence1.TabIndex = 4;
            // 
            // txtSentence2
            // 
            this.txtSentence2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSentence2.Location = new System.Drawing.Point(354, 163);
            this.txtSentence2.Name = "txtSentence2";
            this.txtSentence2.Size = new System.Drawing.Size(418, 20);
            this.txtSentence2.TabIndex = 5;
            // 
            // txtSentence3
            // 
            this.txtSentence3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSentence3.Location = new System.Drawing.Point(354, 207);
            this.txtSentence3.Name = "txtSentence3";
            this.txtSentence3.Size = new System.Drawing.Size(418, 20);
            this.txtSentence3.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(333, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Word";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(578, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Translate";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(282, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Ornek 1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(282, 170);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Ornek 2";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(282, 210);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(45, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Ornek 3";
            // 
            // shwButton
            // 
            this.shwButton.Location = new System.Drawing.Point(534, 84);
            this.shwButton.Name = "shwButton";
            this.shwButton.Size = new System.Drawing.Size(130, 22);
            this.shwButton.TabIndex = 12;
            this.shwButton.Text = "Show";
            this.shwButton.UseVisualStyleBackColor = true;
            this.shwButton.Click += new System.EventHandler(this.shwButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.shwButton);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtSentence3);
            this.Controls.Add(this.txtSentence2);
            this.Controls.Add(this.txtSentence1);
            this.Controls.Add(this.txtTranslation);
            this.Controls.Add(this.txtWord);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnNext);
            this.Name = "Form1";
            this.Text = "English Word Teacher";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.TextBox txtWord;
        private System.Windows.Forms.TextBox txtTranslation;
        private System.Windows.Forms.TextBox txtSentence1;
        private System.Windows.Forms.TextBox txtSentence2;
        private System.Windows.Forms.TextBox txtSentence3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button shwButton;
    }
}

