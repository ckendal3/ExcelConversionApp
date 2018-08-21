using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelConversionApp
{
    partial class NewMapControl
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

        /// <summary>
        /// Get the numeric value of the import cell to map from.
        /// </summary>
        /// <returns>int</returns>
        public int GetImportID()
        {
            return Convert.ToInt32(input_ImportID.Text);
        }

        /// <summary>
        /// Get the numeric value of the cell to map the data to be exported to.
        /// </summary>
        /// <returns>int</returns>
        public int GetExportID()
        {
            return Convert.ToInt32(input_ExportID.Text);
        }

        /// <summary>
        /// Gets the name of the mapping. Purely for the user.
        /// </summary>
        /// <returns>string name</returns>
        public string GetCellMappingName()
        {
            return input_MapName.Text;
        }

        /// <summary>
        /// Clears all input fields on the control.
        /// </summary>
        public void ClearControl()
        {
            input_ImportID.Clear();
            input_ExportID.Clear();
            input_MapName.Clear();
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.input_ImportID = new System.Windows.Forms.MaskedTextBox();
            this.input_ExportID = new System.Windows.Forms.MaskedTextBox();
            this.input_MapName = new System.Windows.Forms.TextBox();
            this.label_ImportID = new System.Windows.Forms.Label();
            this.label_MapName = new System.Windows.Forms.Label();
            this.label_ExportID = new System.Windows.Forms.Label();
            this.button_AddMap = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // input_ImportID
            // 
            this.input_ImportID.Location = new System.Drawing.Point(3, 21);
            this.input_ImportID.Mask = "00000";
            this.input_ImportID.Name = "input_ImportID";
            this.input_ImportID.Size = new System.Drawing.Size(100, 20);
            this.input_ImportID.TabIndex = 1;
            this.input_ImportID.ValidatingType = typeof(int);
            // 
            // input_ExportID
            // 
            this.input_ExportID.Location = new System.Drawing.Point(215, 21);
            this.input_ExportID.Mask = "00000";
            this.input_ExportID.Name = "input_ExportID";
            this.input_ExportID.Size = new System.Drawing.Size(100, 20);
            this.input_ExportID.TabIndex = 5;
            this.input_ExportID.ValidatingType = typeof(int);
            // 
            // input_MapName
            // 
            this.input_MapName.Location = new System.Drawing.Point(109, 21);
            this.input_MapName.Name = "input_MapName";
            this.input_MapName.Size = new System.Drawing.Size(100, 20);
            this.input_MapName.TabIndex = 6;
            // 
            // label_ImportID
            // 
            this.label_ImportID.AutoSize = true;
            this.label_ImportID.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_ImportID.Location = new System.Drawing.Point(0, 0);
            this.label_ImportID.Name = "label_ImportID";
            this.label_ImportID.Size = new System.Drawing.Size(68, 18);
            this.label_ImportID.TabIndex = 7;
            this.label_ImportID.Text = "Import ID";
            // 
            // label_MapName
            // 
            this.label_MapName.AutoSize = true;
            this.label_MapName.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_MapName.Location = new System.Drawing.Point(106, 0);
            this.label_MapName.Name = "label_MapName";
            this.label_MapName.Size = new System.Drawing.Size(81, 18);
            this.label_MapName.TabIndex = 8;
            this.label_MapName.Text = "Map Name";
            // 
            // label_ExportID
            // 
            this.label_ExportID.AutoSize = true;
            this.label_ExportID.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_ExportID.Location = new System.Drawing.Point(212, 0);
            this.label_ExportID.Name = "label_ExportID";
            this.label_ExportID.Size = new System.Drawing.Size(69, 18);
            this.label_ExportID.TabIndex = 9;
            this.label_ExportID.Text = "Export ID";
            // 
            // button_AddMap
            // 
            this.button_AddMap.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_AddMap.Location = new System.Drawing.Point(321, 7);
            this.button_AddMap.Name = "button_AddMap";
            this.button_AddMap.Size = new System.Drawing.Size(92, 44);
            this.button_AddMap.TabIndex = 10;
            this.button_AddMap.Text = "Add Map";
            this.button_AddMap.UseVisualStyleBackColor = true;
            // 
            // NewMapControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button_AddMap);
            this.Controls.Add(this.label_ExportID);
            this.Controls.Add(this.label_MapName);
            this.Controls.Add(this.label_ImportID);
            this.Controls.Add(this.input_MapName);
            this.Controls.Add(this.input_ExportID);
            this.Controls.Add(this.input_ImportID);
            this.Name = "NewMapControl";
            this.Size = new System.Drawing.Size(418, 54);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MaskedTextBox input_ImportID;
        private System.Windows.Forms.MaskedTextBox input_ExportID;
        private System.Windows.Forms.TextBox input_MapName;
        private System.Windows.Forms.Label label_ImportID;
        private System.Windows.Forms.Label label_MapName;
        private System.Windows.Forms.Label label_ExportID;
        private System.Windows.Forms.Button button_AddMap;
    }
}
