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
	[Description("Get the mails via Exchange Server.")]
	public class GetExchangeMails : CodeActivity
	{
		[Category("Input")]
		[RequiredArgument]
		[Description("Exchange server's url address. It can be https://outlook.office365.com/EWS/Exchange.asmx for Office365 by default")]
		public InArgument<String> ExchangeServerUrl { get; set; }

		[Category("Input")]
		[RequiredArgument]
		[Description("Exchange server version.")]
		public InArgument<ExchangeVersion> ExchangeVersion { get; set; }

		[Category("Input")]
		[RequiredArgument]
		[Description("Email address of the Exchange account.")]
		public InArgument<String> ExchangeEmailAddress { get; set; }

		[Category("Input")]
		[RequiredArgument]
		[Description("Password of the Exchange account.")]
		public InArgument<String> ExchangePassword { get; set; }

		[Category("Input")]
		[RequiredArgument]
		[Description("Maximum number of e-mails to be read. If you want to read all mails, you must enter a value of -1.")]
		public InArgument<int> Top { get; set; }

		[Category("Output")]
		[Description("The list to which the read e-mail objects will be added.")]
		public OutArgument<List<EmailMessage>> MailMessages { get; set; }

		protected override void Execute(CodeActivityContext context)
		{
			try
			{
				Console.WriteLine("GetExchangeMails activity started.");
				int top = Top.Get(context);
				var service = new ExchangeService(ExchangeVersion.Get(context))
				{
					Credentials = new NetworkCredential(ExchangeEmailAddress.Get(context), ExchangePassword.Get(context), ""),
					Url = new Uri(ExchangeServerUrl.Get(context))
				};

				var inbox = Folder.Bind(service, WellKnownFolderName.Inbox);

				if (top == -1)
				{
					top = inbox.TotalCount;
				}
				else if (top < 0)
				{
					top = 0;
				}

				if (top > 0)
				{
					var view = new ItemView(top) { PropertySet = PropertySet.FirstClassProperties };
					List<EmailMessage> emailMessages = service.FindItems(inbox.Id, view).Items.Select(x => EmailMessage.Bind(service, new ItemId(x.Id.UniqueId.ToString()))).ToList();
					MailMessages.Set(context, emailMessages);
					Console.WriteLine("GetExchangeMails aktivitesinde mailler başarıyla okundu. Okunan mail sayısı: " + MailMessages.Get(context).Count.ToString());
				}
				else
				{
					MailMessages.Set(context, new List<EmailMessage>());
					Console.WriteLine("Herhangi bir mail bulunamadı.");
				}
				Console.WriteLine("GetExchangeMails aktivitesi başarıyla tamamlandı.");
			}
			catch (Exception e)
			{
				Console.WriteLine("Mailler okunurken bir hata oluştu. Credentials bilgileri yanlış olabilir. Hata mesajı: " + e.Message);
				throw new Exception("Mailler okunurken bir hata oluştu. Credentials bilgileri yanlış olabilir. Hata mesajı: "+ e.Message);
			}
		}
	}
}
