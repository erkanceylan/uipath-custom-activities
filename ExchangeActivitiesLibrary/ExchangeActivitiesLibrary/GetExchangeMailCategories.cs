using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.IO;

namespace ExchangeActivities
{
	[Description("Returns the categories of the mail.")]
	public class GetExchangeMailCategories : CodeActivity
	{
		[Category("Input")]
		[RequiredArgument]
		[Description("The e-mail object whose categories are to be read.")]
		public InArgument<EmailMessage> MailMessage { get; set; }

		[Category("Output")]
		[Description("The list to which the categories of the email will be added.")]
		public OutArgument<List<String>> MailCategories { get; set; }

		protected override void Execute(CodeActivityContext context)
		{
			Console.WriteLine("GetExchangeMailCategories activity started.");

			try
			{
				MailCategories.Set(context, new List<String>());
				List<String> categories = MailMessage.Get(context).Categories.ToList();
				MailCategories.Set(context, categories);
				Console.WriteLine("All categories of the mail have been read. Categories: " + String.Join(" | ", categories).ToString());

				Console.WriteLine("GetExchangeMailCategories activity completed successfully.");
			}
			catch (Exception e)
			{
				Console.WriteLine("There was an error reading the mail categories. Error message: " + e.Message);
				throw new Exception("There was an error reading the mail categories. Error message: " + e.Message);
			}
		}
	}
}
