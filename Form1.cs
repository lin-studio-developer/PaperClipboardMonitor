using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ClipboardMonitor
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		[DllImport("User32.dll")]
		protected static extern int SetClipboardViewer(int hWndNewViewer);

		[DllImport("User32.dll", CharSet=CharSet.Auto)]
		public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

		[DllImport("user32.dll", CharSet=CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

		IntPtr nextClipboardViewer;

        private Regex regexRemoveLineBreak = new Regex(@"\s*(\\pard|\\par)\s*", RegexOptions.Compiled);
        private Regex regexQuote = new Regex(@"“|”", RegexOptions.Compiled); //
        private Regex regexRemoveSuperSubScript = new Regex(@"\{\\fs14\s+\d+\s+\}", RegexOptions.Compiled);
        private Regex regexRemoveReference = new Regex(@"\s*\[(\d|\-|\,|\s|–)+\]\s*", RegexOptions.Compiled);
        private Regex regexRemoveReferenceForDetectEndOfSentence = new Regex(@"\s*\[(\d|\-|\,|\s|–)+\]\s*(?<ch>(\,|\.){1})", RegexOptions.Compiled);

        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private CheckBox checkBox3;
        private CheckBox checkBox4;
        private NotifyIcon notifyIcon1;
        private CheckBox chkRemoveReferences;
        private IContainer components;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			nextClipboardViewer = (IntPtr)SetClipboardViewer((int) this.Handle);

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

        public static string matchEvaluatorMethod(Match match)
        {
            return match.Groups["ch"].Value;
        }

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			ChangeClipboardChain(this.Handle, nextClipboardViewer);
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.chkRemoveReferences = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Enabled = false;
            this.checkBox1.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.checkBox1.Location = new System.Drawing.Point(12, 12);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(358, 17);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "Auto Complete break-line word (line1: privi-, line2: lege => privilege)";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Checked = true;
            this.checkBox2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox2.Enabled = false;
            this.checkBox2.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.checkBox2.Location = new System.Drawing.Point(12, 34);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(449, 17);
            this.checkBox2.TabIndex = 1;
            this.checkBox2.Text = "Auto help generate a \"complete\" and \"plain-text\" sentence (No need to paste as pl" +
    "ain-text)";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Checked = true;
            this.checkBox3.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox3.Enabled = false;
            this.checkBox3.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.checkBox3.Location = new System.Drawing.Point(12, 56);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(228, 17);
            this.checkBox3.TabIndex = 2;
            this.checkBox3.Text = "Auto help remove Subscript and Supscript";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Checked = true;
            this.checkBox4.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox4.Enabled = false;
            this.checkBox4.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.checkBox4.Location = new System.Drawing.Point(12, 79);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(309, 17);
            this.checkBox4.TabIndex = 3;
            this.checkBox4.Text = "Auto help transform Chinese “quote” to English \"quote\"";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseClick);
            // 
            // chkRemoveReferences
            // 
            this.chkRemoveReferences.AutoSize = true;
            this.chkRemoveReferences.Checked = true;
            this.chkRemoveReferences.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRemoveReferences.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.chkRemoveReferences.Location = new System.Drawing.Point(12, 102);
            this.chkRemoveReferences.Name = "chkRemoveReferences";
            this.chkRemoveReferences.Size = new System.Drawing.Size(370, 17);
            this.chkRemoveReferences.TabIndex = 4;
            this.chkRemoveReferences.Text = "Auto help remove reference number (ex: [12]  [23, 24]  [1,2,5]  [45-47])";
            this.chkRemoveReferences.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 15);
            this.ClientSize = new System.Drawing.Size(814, 153);
            this.Controls.Add(this.chkRemoveReferences);
            this.Controls.Add(this.checkBox4);
            this.Controls.Add(this.checkBox3);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Clipboard Optimization for Papers (by Miles Lin)";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		protected override void WndProc(ref System.Windows.Forms.Message m)
		{
			// defined in winuser.h
			const int WM_DRAWCLIPBOARD = 0x308;
			const int WM_CHANGECBCHAIN = 0x030D;

            try
            {

                switch (m.Msg)
                {
                    case WM_DRAWCLIPBOARD:
                        DisplayClipboardData();
                        SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                        break;

                    case WM_CHANGECBCHAIN:
                        if (m.WParam == nextClipboardViewer)
                            nextClipboardViewer = m.LParam;
                        else
                            SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                        break;

                    default:
                        base.WndProc(ref m);
                        break;
                }

            }
            catch
            {

            }
		}

        string handleAllString(ref string strPlainText)
        {
            
            if (chkRemoveReferences.Checked)
            {
                strPlainText = regexRemoveReferenceForDetectEndOfSentence.Replace(strPlainText, matchEvaluatorMethod);

                strPlainText = regexRemoveReference.Replace(strPlainText, " ");
            }

            return strPlainText;
        }

        string convertToPlainText(string strRichText)
        {
            string strPlainText = strRichText;
            try
            {
                using (RichTextBox rtb = new RichTextBox())
                {
                    handleAllString(ref strRichText);

                    string strModifiedText = regexRemoveLineBreak.Replace(regexRemoveSuperSubScript.Replace(strRichText, " "), " ");

                    rtb.Rtf = strModifiedText;

                    strPlainText = rtb.Text;

                    strPlainText = regexQuote.Replace(strPlainText, "\"");
                }
            }
            catch
            {
                
            }
            return strPlainText.Trim();
        }

		void DisplayClipboardData()		
		{
			try
			{
				IDataObject iData = new DataObject();  
				iData = Clipboard.GetDataObject();

                if (iData.GetDataPresent(DataFormats.Rtf))
                {
                    string dtData = iData.GetData(DataFormats.Rtf) as string;

                    if (!string.IsNullOrEmpty(dtData))
                    {
                        Clipboard.SetText(convertToPlainText(dtData));
                    }
                }
                else if (iData.GetDataPresent(DataFormats.Text))
                {
                    string dtData = iData.GetData(DataFormats.Text) as string;

                    if (!string.IsNullOrEmpty(dtData))
                    {

                        dtData = regexQuote.Replace(dtData, "\"");

                        handleAllString(ref dtData);

                        Clipboard.SetText(dtData.Replace(Environment.NewLine, ""));
                    }
                }
			}
			catch
			{
				
			}
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            this.Show();
            notifyIcon1.Visible=false;
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon1.Visible = true;
                notifyIcon1.ShowBalloonTip(7, this.Text, "This program is now running on the background, you can restore the window state by clicking the icon on the taskbar.", ToolTipIcon.Info);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Text = this.Text;
            //
        }

	}
}
