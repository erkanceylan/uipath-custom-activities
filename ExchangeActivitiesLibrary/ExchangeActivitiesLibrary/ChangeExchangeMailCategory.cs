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
	[Description("Changes the category of the mail.")]
	public class ChangeExchangeMailCategory : CodeActivity
	{
		[Category("Input")]
		[RequiredArgument]
		[Description("Email object to change category information.")]
		public InArgument<EmailMessage> MailMessage { get; set; }

		[Category("Input")]
		[RequiredArgument]
		[Description("Current category information of the email to be changed.")]
		public InArgument<String> CurrentCategory { get; set; }

		[Category("Input")]
		[RequiredArgument]
		[Description("The new category information to be added to the email.")]
		public InArgument<String> NewCategory { get; set; }

		protected override void Execute(CodeActivityContext context)
		{
			Console.WriteLine("ChangeExchangeMailCategory aktivitesi çalışmaya başladı.");

			try
			{
				var categories = MailMessage.Get(context).Categories;
				if (categories.Contains(CurrentCategory.Get(context)))
				{
					MailMessage.Get(context).Categories.Remove(CurrentCategory.Get(context));
					Console.WriteLine("The current category of the mail has been deleted.");
				}
				else
				{
					Console.WriteLine("Deletion is not performed because the searched category does not exist in the mail.");
				}

				if (categories.Contains(NewCategory.Get(context)))
				{
					Console.WriteLine("The category to add already exists.");
				}
				else
				{
					MailMessage.Get(context).Categories.Add(NewCategory.Get(context));
					Console.WriteLine("New category added to the mail.");
				}

				MailMessage.Get(context).Update(ConflictResolutionMode.AutoResolve);
				Console.WriteLine("The category of the mail has been successfully changed. Old category: " + CurrentCategory.Get(context) + ", new category: " + NewCategory.Get(context));
				Console.WriteLine("ChangeExchangeMailCategory activity completed successfully.");
			}
			catch (Exception e)
			{
				Console.WriteLine("An error occurred while changing the category. Error message:: " + e.Message);
				throw new Exception("An error occurred while changing the category. Error message:: " + e.Message);
			}
		}
	}

}