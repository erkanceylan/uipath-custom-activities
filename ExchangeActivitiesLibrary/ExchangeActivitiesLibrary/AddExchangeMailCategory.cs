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
	[Description("Adds a new category to the e-mail.")]
	public class AddExchangeMailCategory : CodeActivity
	{
		[Category("Input")]
		[RequiredArgument]
		[Description("Email object to add category information.")]
		public InArgument<EmailMessage> MailMessage { get; set; }

		[Category("Input")]
		[RequiredArgument]
		[Description("The category information to be attached to the email.")]
		public InArgument<String> Category { get; set; }

		protected override void Execute(CodeActivityContext context)
		{
			Console.WriteLine("AddExchangeMailCategory activity started.");

			try
			{
				var categories = MailMessage.Get(context).Categories;
				if (categories.Contains(Category.Get(context)))
				{
					Console.WriteLine("The category to add already exists.");
				}
				else
				{
					MailMessage.Get(context).Categories.Add(Category.Get(context));
					Console.WriteLine("New category added to the mail.");
				}
				MailMessage.Get(context).Update(ConflictResolutionMode.AutoResolve);
				Console.WriteLine("New category(" + Category.Get(context) + ") has been successfully added to the mail and saved.");
				Console.WriteLine("AddExchangeMailCategory activity completed successfully.");
			}
			catch (Exception e)
			{
				Console.WriteLine("An error occurred while adding the category. Error message: " + e.Message);
				throw new Exception("An error occurred while adding the category. Error message: " + e.Message);
			}
		}
	}

}
