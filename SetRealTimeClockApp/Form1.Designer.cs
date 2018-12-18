namespace SetRealTimeClockApp
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.setClock_button = new System.Windows.Forms.Button();
            this.serialPort = new System.IO.Ports.SerialPort(this.components);
            this.SuspendLayout();
            // 
            // setClock_button
            // 
            this.setClock_button.Location = new System.Drawing.Point(73, 93);
            this.setClock_button.Name = "setClock_button";
            this.setClock_button.Size = new System.Drawing.Size(75, 23);
            this.setClock_button.TabIndex = 1;
            this.setClock_button.Text = "時計セット";
            this.setClock_button.UseVisualStyleBackColor = true;
            this.setClock_button.Click += new System.EventHandler(this.setClock_button_Click);
            // 
            // serialPort
            // 
            this.serialPort.DtrEnable = true;
            this.serialPort.PortName = "COM5";
            this.serialPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.serialPort_DataReceived);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(239, 127);
            this.ControlBox = false;
            this.Controls.Add(this.setClock_button);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "時刻設定";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button setClock_button;
        private System.IO.Ports.SerialPort serialPort;
    }
}

