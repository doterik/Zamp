using System;
using System.Windows.Forms;
using System.Drawing;

namespace HelloWorld
{
	public partial class ADXOlForm1 : AddinExpress.OL.ADXOlForm
	{
		public ADXOlForm1()
		{
			InitializeComponent();
		}

		private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			this.linkLabel1.LinkVisited = true;
			System.Diagnostics.Process.Start("http://www.add-in-express.com/outlook-extension/");
		}
	}
}
