namespace ConfigEditor
{
    partial class Configurator
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Configurator));
            this.btSaveAs = new System.Windows.Forms.Button();
            this.listView = new System.Windows.Forms.ListView();
            this.clValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.clType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.clName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btOpen = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btSaveAs
            // 
            this.btSaveAs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btSaveAs.Location = new System.Drawing.Point(486, 8);
            this.btSaveAs.Name = "btSaveAs";
            this.btSaveAs.Size = new System.Drawing.Size(93, 23);
            this.btSaveAs.TabIndex = 0;
            this.btSaveAs.Text = "Сохранить как";
            this.btSaveAs.UseVisualStyleBackColor = true;
            this.btSaveAs.Click += new System.EventHandler(this.btSaveAs_Click);
            // 
            // listView
            // 
            this.listView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.clValue,
            this.clType,
            this.clName});
            this.listView.FullRowSelect = true;
            this.listView.GridLines = true;
            this.listView.HideSelection = false;
            this.listView.LabelEdit = true;
            this.listView.Location = new System.Drawing.Point(12, 37);
            this.listView.Name = "listView";
            this.listView.Size = new System.Drawing.Size(567, 432);
            this.listView.TabIndex = 1;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            this.listView.AfterLabelEdit += new System.Windows.Forms.LabelEditEventHandler(this.listView_AfterLabelEdit);
            this.listView.DoubleClick += new System.EventHandler(this.listView_DoubleClick);
            // 
            // clValue
            // 
            this.clValue.Text = "Value";
            this.clValue.Width = 117;
            // 
            // clType
            // 
            this.clType.Text = "Type";
            this.clType.Width = 147;
            // 
            // clName
            // 
            this.clName.Text = "Name";
            this.clName.Width = 271;
            // 
            // btOpen
            // 
            this.btOpen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btOpen.Location = new System.Drawing.Point(387, 8);
            this.btOpen.Name = "btOpen";
            this.btOpen.Size = new System.Drawing.Size(93, 23);
            this.btOpen.TabIndex = 0;
            this.btOpen.Text = "Открыть";
            this.btOpen.UseVisualStyleBackColor = true;
            this.btOpen.Click += new System.EventHandler(this.btOpen_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "Config.dll";
            this.openFileDialog.Filter = "Library (*.dll)|*.dll";
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.FileName = "Config.dll";
            this.saveFileDialog.Filter = "Library (*.dll)|*.dll";
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(57, 10);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(202, 20);
            this.textBox1.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Метка";
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.Location = new System.Drawing.Point(315, 10);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(66, 20);
            this.textBox2.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(265, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Версия";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // Configurator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 481);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btOpen);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.btSaveAs);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Configurator";
            this.Text = "Конфигуратор";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btSaveAs;
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.ColumnHeader clValue;
        private System.Windows.Forms.ColumnHeader clType;
        private System.Windows.Forms.ColumnHeader clName;
        private System.Windows.Forms.Button btOpen;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;

    }
}

