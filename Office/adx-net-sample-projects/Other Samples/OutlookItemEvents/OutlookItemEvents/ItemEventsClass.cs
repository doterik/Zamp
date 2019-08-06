using System;
using System.Reflection;
using System.Windows.Forms;

namespace OutlookItemEvents
{
	/// <summary>
	/// Add-in Express Outlook Item Events Class
	/// </summary>
	public class ItemEventsClass : AddinExpress.MSO.ADXOutlookItemEvents
	{
		public ItemEventsClass(AddinExpress.MSO.ADXAddinModule module)
			: base(module)
		{
		}

		public override void ProcessAttachmentAdd(object attachment)
		{
			// TODO: Add some code
		}

		public override void ProcessAttachmentRead(object attachment)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeAttachmentSave(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeCheckNames(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessClose(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			if (this.IsConnected)
			{
				Outlook.OlObjectClass _class = (Outlook.OlObjectClass)Convert.ToInt32(
					ItemObj.GetType().InvokeMember("Class", BindingFlags.GetProperty, null, ItemObj, null));
				if (MessageBox.Show("Do you really want to close the " + _class.ToString() + " item?",
					"Outlook Item Events Example", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk) == DialogResult.Cancel)
				{
					e.Cancel = true;
				}
			}
		}

		public override void ProcessCustomAction(object action, object response, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessCustomPropertyChange(string name)
		{
			// TODO: Add some code
		}

		public override void ProcessForward(object forward, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessOpen(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessPropertyChange(string name)
		{
			// TODO: Add some code
		}

		public override void ProcessRead()
		{
			// TODO: Add some code
		}

		public override void ProcessReply(object response, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessReplyAll(object response, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessSend(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessWrite(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeDelete(object item, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessAttachmentRemove(object attachment)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeAttachmentAdd(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeAttachmentPreview(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeAttachmentRead(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeAttachmentWriteToTempFile(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessUnload()
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeAutoSave(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			// TODO: Add some code
		}

		public override void ProcessBeforeRead()
		{
			// TODO: Add some code
		}

		public override void ProcessAfterWrite()
		{
			// TODO: Add some code
		}
	}
}
