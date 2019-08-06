using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SimpleApp
{
	public enum OlDefaultFolders
	{
		olFolderDeletedItems = 3,
		olFolderOutbox = 4,
		olFolderSentMail = 5,
		olFolderInbox = 6,
		olFolderCalendar = 9,
		olFolderContacts = 10,
		olFolderJournal = 11,
		olFolderNotes = 12,
		olFolderTasks = 13,
		olFolderDrafts = 16,
		olPublicFoldersAllPublicFolders = 18,
		olFolderConflicts = 19,
		olFolderSyncIssues = 20,
		olFolderLocalFailures = 21,
		olFolderServerFailures = 22,
		olFolderJunk = 23
	}

	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.RadioButton rbON;
		private System.Windows.Forms.RadioButton rbOFF;
		private System.Windows.Forms.RadioButton rbOOM;
		private System.Windows.Forms.RadioButton rbCDO;
		private System.Windows.Forms.RadioButton rbMAPI;
		private System.Windows.Forms.Label lbInfo;
		private System.Windows.Forms.Button btnGetInfo;
		private System.Windows.Forms.RichTextBox tbInfo;
		private AddinExpress.Outlook.SecurityManager securityManager1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			InitializeComponent();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.rbMAPI = new System.Windows.Forms.RadioButton();
			this.rbCDO = new System.Windows.Forms.RadioButton();
			this.rbOOM = new System.Windows.Forms.RadioButton();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.rbOFF = new System.Windows.Forms.RadioButton();
			this.rbON = new System.Windows.Forms.RadioButton();
			this.lbInfo = new System.Windows.Forms.Label();
			this.btnGetInfo = new System.Windows.Forms.Button();
			this.tbInfo = new System.Windows.Forms.RichTextBox();
			this.securityManager1 = new AddinExpress.Outlook.SecurityManager();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.rbMAPI);
			this.groupBox1.Controls.Add(this.rbCDO);
			this.groupBox1.Controls.Add(this.rbOOM);
			this.groupBox1.Location = new System.Drawing.Point(16, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(176, 112);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Access Type";
			// 
			// rbMAPI
			// 
			this.rbMAPI.Location = new System.Drawing.Point(16, 72);
			this.rbMAPI.Name = "rbMAPI";
			this.rbMAPI.Size = new System.Drawing.Size(104, 24);
			this.rbMAPI.TabIndex = 2;
			this.rbMAPI.Text = "Simple MAPI";
			// 
			// rbCDO
			// 
			this.rbCDO.Location = new System.Drawing.Point(16, 48);
			this.rbCDO.Name = "rbCDO";
			this.rbCDO.Size = new System.Drawing.Size(104, 24);
			this.rbCDO.TabIndex = 1;
			this.rbCDO.Text = "CDO";
			// 
			// rbOOM
			// 
			this.rbOOM.Checked = true;
			this.rbOOM.Location = new System.Drawing.Point(16, 24);
			this.rbOOM.Name = "rbOOM";
			this.rbOOM.Size = new System.Drawing.Size(104, 24);
			this.rbOOM.TabIndex = 0;
			this.rbOOM.TabStop = true;
			this.rbOOM.Text = "OOM";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.rbOFF);
			this.groupBox2.Controls.Add(this.rbON);
			this.groupBox2.Location = new System.Drawing.Point(208, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(152, 88);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Security";
			// 
			// rbOFF
			// 
			this.rbOFF.Location = new System.Drawing.Point(24, 48);
			this.rbOFF.Name = "rbOFF";
			this.rbOFF.Size = new System.Drawing.Size(104, 24);
			this.rbOFF.TabIndex = 1;
			this.rbOFF.Text = "OFF";
			// 
			// rbON
			// 
			this.rbON.Checked = true;
			this.rbON.Location = new System.Drawing.Point(24, 24);
			this.rbON.Name = "rbON";
			this.rbON.Size = new System.Drawing.Size(104, 24);
			this.rbON.TabIndex = 0;
			this.rbON.TabStop = true;
			this.rbON.Text = "ON";
			// 
			// lbInfo
			// 
			this.lbInfo.Location = new System.Drawing.Point(3, 147);
			this.lbInfo.Name = "lbInfo";
			this.lbInfo.Size = new System.Drawing.Size(352, 24);
			this.lbInfo.TabIndex = 2;
			this.lbInfo.Text = "Information:";
			this.lbInfo.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			// 
			// btnGetInfo
			// 
			this.btnGetInfo.Location = new System.Drawing.Point(280, 112);
			this.btnGetInfo.Name = "btnGetInfo";
			this.btnGetInfo.Size = new System.Drawing.Size(75, 23);
			this.btnGetInfo.TabIndex = 3;
			this.btnGetInfo.Text = "Get Info";
			this.btnGetInfo.Click += new System.EventHandler(this.btnGetInfo_Click);
			// 
			// tbInfo
			// 
			this.tbInfo.BackColor = System.Drawing.SystemColors.Window;
			this.tbInfo.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.tbInfo.Location = new System.Drawing.Point(0, 176);
			this.tbInfo.Name = "tbInfo";
			this.tbInfo.ReadOnly = true;
			this.tbInfo.Size = new System.Drawing.Size(370, 232);
			this.tbInfo.TabIndex = 4;
			this.tbInfo.Text = "";
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(370, 408);
			this.Controls.Add(this.tbInfo);
			this.Controls.Add(this.btnGetInfo);
			this.Controls.Add(this.lbInfo);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Name = "Form1";
			this.Text = "Simple App";
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

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

		private void btnGetInfo_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			if (rbOOM.Checked)
				DoOOM();
			if (rbCDO.Checked)
				DoCDO();
			if (rbMAPI.Checked)
				DoSimpleMAPI();
			this.Cursor = Cursors.Default;
		}

		private void DoOOM()
		{
			object nameSpace = null, items = null, folder = null;
			object outlookApp = null, mailItem = null;
			tbInfo.Clear();
			Type olType = Type.GetTypeFromProgID("Outlook.Application", false);
			if (olType != null)
			{
				try
				{
					outlookApp = Activator.CreateInstance(olType);
					if (outlookApp != null)
					{
						if (rbOFF.Checked)
						{
							securityManager1.ConnectTo(outlookApp);
							// switch OFF
							securityManager1.DisableOOMWarnings = true;
						}
						try
						{
							nameSpace = outlookApp.GetType().InvokeMember("GetNamespace", BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });
							folder = nameSpace.GetType().InvokeMember("GetDefaultFolder", BindingFlags.InvokeMethod, null, nameSpace, new object[] { OlDefaultFolders.olFolderInbox });
							items = folder.GetType().InvokeMember("Items", BindingFlags.GetProperty, null, folder, null);
							mailItem = items.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, items, new object[] { 1 });
							if (mailItem != null)
							{
								string[] lines = new string[4];
								lines[0] = ("*** First message in Microsoft Outlook inbox (Outlook Object Model) ***");
								lines[1] = "";
								lines[2] = "From: " + Convert.ToString(mailItem.GetType().InvokeMember("SenderName", BindingFlags.GetProperty, null, mailItem, null));
								lines[3] = "Subject: " + Convert.ToString(mailItem.GetType().InvokeMember("Subject", BindingFlags.GetProperty, null, mailItem, null));
								tbInfo.Lines = lines;
								Marshal.ReleaseComObject(mailItem);
							}
						}
						catch { }
					}
				}
				finally
				{
					if (items != null)
						Marshal.ReleaseComObject(items);
					if (folder != null)
						Marshal.ReleaseComObject(folder);
					if (nameSpace != null)
						Marshal.ReleaseComObject(nameSpace);
					if (rbOFF.Checked)
					{
						// switch ON
						securityManager1.DisableOOMWarnings = false;
						securityManager1.Disconnect(outlookApp);
					}
					if (outlookApp != null)
						Marshal.ReleaseComObject(outlookApp);
				}
			}
		}

		private void DoCDO()
		{
			object mapiSession = null, objItem = null, inbox = null, messages = null, sender = null;
			tbInfo.Clear();
			Type tpMapi = Type.GetTypeFromProgID("MAPI.Session", false);
			if (tpMapi != null)
			{
				try
				{
					mapiSession = Activator.CreateInstance(tpMapi);
					if (mapiSession != null)
					{
						try
						{
							mapiSession.GetType().InvokeMember("Logon", BindingFlags.InvokeMethod, null, mapiSession, null);

							if (rbOFF.Checked)
								securityManager1.DisableCDOWarnings = true;

							inbox = mapiSession.GetType().InvokeMember("Inbox", BindingFlags.GetProperty, null, mapiSession, null);
							messages = inbox.GetType().InvokeMember("Messages", BindingFlags.GetProperty, null, inbox, null);
							objItem = messages.GetType().InvokeMember("GetFirst", BindingFlags.GetProperty, null, messages, null);
							if (objItem != null)
							{
								string[] lines = new string[4];
								lines[0] = ("*** First message in Microsoft Outlook inbox (Outlook Object Model) ***");
								lines[1] = "";
								sender = objItem.GetType().InvokeMember("Sender", BindingFlags.GetProperty, null, objItem, null);
								lines[2] = "From: " + Convert.ToString(sender.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, sender, null));
								lines[3] = "Subject: " + Convert.ToString(objItem.GetType().InvokeMember("Subject", BindingFlags.GetProperty, null, objItem, null));
								tbInfo.Lines = lines;
							}
						}
						catch { }
					}
				}
				finally
				{
					if (rbOFF.Checked)
						securityManager1.DisableCDOWarnings = false;
					if (sender != null)
						Marshal.ReleaseComObject(sender);
					if (objItem != null)
						Marshal.ReleaseComObject(objItem);
					if (messages != null)
						Marshal.ReleaseComObject(messages);
					if (inbox != null)
						Marshal.ReleaseComObject(inbox);
					if (mapiSession != null)
						Marshal.ReleaseComObject(mapiSession);
				}
			}
		}

		[DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
		public static extern int MAPILogon(int ulUIParam,
			[MarshalAs(UnmanagedType.LPWStr)] string lpszProfileName,
			[MarshalAs(UnmanagedType.LPWStr)] string lpszPassword,
			int flFlags, int ulReserved, out IntPtr lplhSession);

		[DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
		public static extern int MAPILogoff(IntPtr lhSession, int ulUIParam, int flFlags, int ulReserved);

		[DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
		public static extern int MAPIFindNext(IntPtr lhSession, int ulUIParam,
			[MarshalAs(UnmanagedType.LPWStr)] string lpszMessageType,
			[MarshalAs(UnmanagedType.LPWStr)] string lpszSeedMessageID,
			int flFlags, int ulReserved, IntPtr lpszMessageID);

		[DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
		public static extern int MAPIReadMail(IntPtr lhSession, int ulUIParam,
			IntPtr lpszMessageID, int flFlags, int ulReserved, ref IntPtr lppMessage);

		[DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
		public static extern int MAPIFreeBuffer(IntPtr lpBuffer);

		private const int MAPI_LOGON_UI = 0x00000001;
		private const int MAPI_LONG_MSGID = 0x00004000;
		private const int MAPI_ENVELOPE_ONLY = 0x00000040;
		private const int MAPI_PEEK = 0x00000080;
		private const int MAPI_SUPPRESS_ATTACH = 0x00000800;

		[StructLayout(LayoutKind.Sequential)]
		public struct MapiMessage
		{
			public uint ulReserved;				// Reserved for future use (M.B. 0)
			[MarshalAs(UnmanagedType.LPStr)]
			public string lpszSubject;			// Message Subject
			[MarshalAs(UnmanagedType.LPStr)]
			public string lpszNoteText;			// Message Text
			[MarshalAs(UnmanagedType.LPStr)]
			public string lpszMessageType;		// Message Class
			[MarshalAs(UnmanagedType.LPStr)]
			public string lpszDateReceived;		// in YYYY/MM/DD HH:MM format
			[MarshalAs(UnmanagedType.LPStr)]
			public string lpszConversationID;	// conversation thread ID
			public uint flFlags;				// unread,return receipt
			public IntPtr lpOriginator;			// Originator descriptor
			public uint nRecipCount;			// Number of recipients
			public IntPtr lpRecips;				// Recipient descriptors
			public uint nFileCount;				// # of file attachments
			public IntPtr lpFiles;				// Attachment descriptors
		}

		[StructLayout(LayoutKind.Sequential)]
		public struct MapiRecipDesc
		{
			public uint ulReserved;		// Reserved for future use
			public uint ulRecipClass;	// Recipient class
			// MAPI_TO, MAPI_CC, MAPI_BCC, MAPI_ORIG
			[MarshalAs(UnmanagedType.LPStr)]
			public string lpszName;		// Recipient name
			[MarshalAs(UnmanagedType.LPStr)]
			public string lpszAddress;	// Recipient address (optional)
			public uint ulEIDSize;				// Count in bytes of size of pEntryID
			public IntPtr lpEntryID;			// System-specific recipient reference
		}

		private void DoSimpleMAPI()
		{
			IntPtr session = IntPtr.Zero;
			string messageID = String.Empty;
			MapiMessage message = new MapiMessage();
			MapiRecipDesc recipient = new MapiRecipDesc();
			AddinExpress.Outlook.SecurityManager securityManager2 = new AddinExpress.Outlook.SecurityManager();
			tbInfo.Clear();
			if (MAPILogon(0, "", "", MAPI_LOGON_UI, 0, out session) == 0)
			{
				if (rbOFF.Checked)
					securityManager2.DisableSMAPIWarnings = true;
				try
				{
					IntPtr mID = Marshal.AllocHGlobal(2048);
					if (mID != IntPtr.Zero)
					{
						try
						{
							MAPIFindNext(session, 0, "", "", MAPI_LONG_MSGID, 0, mID);
							IntPtr lpMessage = IntPtr.Zero;
							MAPIReadMail(session, 0, mID, MAPI_ENVELOPE_ONLY | MAPI_PEEK | MAPI_SUPPRESS_ATTACH, 0, ref lpMessage);
							if (lpMessage != IntPtr.Zero)
							{
								message = (MapiMessage)Marshal.PtrToStructure(lpMessage, typeof(MapiMessage));
								recipient = (MapiRecipDesc)Marshal.PtrToStructure(message.lpOriginator, typeof(MapiRecipDesc));
								string[] lines = new string[4];
								lines[0] = ("*** First message in Microsoft Outlook inbox (Outlook Object Model) ***");
								lines[1] = "";
								lines[2] = "From: " + recipient.lpszName;
								lines[3] = "Subject: " + message.lpszSubject;
								tbInfo.Lines = lines;
								MAPIFreeBuffer(lpMessage);
							}
						}
						catch { }
						finally
						{
							Marshal.FreeHGlobal(mID);
						}
					}
				}
				finally
				{
					if (rbOFF.Checked)
						securityManager2.DisableSMAPIWarnings = false;
					MAPILogoff(session, 0, 0, 0);
				}
			}
		}
	}
}
