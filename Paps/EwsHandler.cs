using Microsoft.Exchange.WebServices.Data;
using Paps.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Paps
{
	public class EwsHandler
	{
		private static ExchangeService service; // Instance of Exchange that will query the mailbox

		public EwsHandler()
		{
			CreateConnection();
		}

		/// <summary>
		/// Creates a connection to EWS and assigns to the Service object
		/// </summary>
		public void CreateConnection()
		{
			//// Return if service is already established
			//if (service != null) { return; }

			ExchangeService newservice = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
			{
				TraceEnabled = true,
				Url = new Uri(@"EXCHANGE.ASMX URL"),
				Credentials = new WebCredentials(
						"USERNAME",
						"PASSWORD",
						"DOMAIN")
			};

			service = newservice;
			newservice = null;

			System.Threading.Tasks.Task.Delay(1740000).ContinueWith(t => CreateConnection());

			//StreamingSubscription();
		}

		/// <summary>
		/// Manipulates a enumerable of Email Messages
		/// </summary>
		/// <param name="clientref">Client Reference number, either TCS or DR</param>
		/// <returns>An instance of PAPs with min, max & avg</returns>
		public PAPs ReturnEmailsFromClientReference(string clientref)
		{
			PAPs paps = new PAPs();

			// Loop over emails and create custom PAPSizing object list
			foreach (var item in FindItemsByClientReference(clientref))
			{
				if (item.Attachments[0] == null) { continue; }

				FileAttachment itemAttachment = (FileAttachment)item.Attachments[0];
				PAPSizing pap = new PAPSizing
				{
					DateGenerated = item.DateTimeReceived,
					FileSize = itemAttachment.Size,
					Subject = item.Subject
				};

				paps.Paps.Add(pap);

			}

			// Get Max
			paps.MaxPap = paps.Paps.Max(i => i.FileSize);

			// Get Min
			paps.MinPap = paps.Paps.Min(i => i.FileSize);

			// Get Average
			paps.AveragePap = paps.Paps.Average(i => i.FileSize);

			// Get Count
			paps.PapCount = paps.Paps.Count();


			return paps;
		}

		/// <summary>
		/// Queries EWS to find all email messages relating to the client reference.
		/// This function is to become more generic which is why the manipulation logic has been split out.
		/// </summary>
		/// <param name="clientref">Client Reference number, either TCS or DR</param>
		/// <returns>An enumerable of EmailMessage to pass for manipulatiion.</returns>
		private List<EmailMessage> FindItemsByClientReference(string clientref)
		{
			if (string.IsNullOrEmpty(clientref))
			{
				return null;
			}

			Mailbox mb = new Mailbox("papretention2@stepchange.org");
			var inboxFolderId = new FolderId(WellKnownFolderName.Inbox, mb);


			//		SearchFilter.ContainsSubstring subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject,
			//clientref, ContainmentMode.Substring, ComparisonMode.IgnoreCase);

			SearchFilter subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, clientref);


			ItemView view = new ItemView(50);
			view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
			view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
			view.Traversal = ItemTraversal.Shallow;


			FindItemsResults<Item> findResults = service.FindItems(inboxFolderId, subjectFilter, view);

			List<EmailMessage> emails = new List<EmailMessage>();



			foreach (Item findResultItem in findResults.Items)
			{
				EmailMessage email = EmailMessage.Bind(service, findResultItem.Id, new PropertySet(BasePropertySet.FirstClassProperties));
				emails.Add(email);
			}

			return emails;
		}

		/// <summary>
		/// Queries the entire mailbox 
		/// </summary>
		/// <returns>An instance of PAPMailboxDefaults which shows the min, max & avg email attachment size.</returns>
		public PAPMailboxDefaults ReturnPapMailboxDefaults()
		{
			PAPMailboxDefaults papDefaults = new PAPMailboxDefaults();
			List<Double> papAverages = new List<double>();
			int offset = 0;
			int pageSize = 50;
			bool moreItems = true;
			Mailbox mb = new Mailbox("papretention2@stepchange.org");
			var inboxFolderId = new FolderId(WellKnownFolderName.Inbox, mb);

			FindItemsResults<Item> findResults;
			ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

			while (moreItems)
			{
				System.Threading.Tasks.Task.Delay(100000).ContinueWith(t => CreateConnection());

				view.PropertySet = BasePropertySet.FirstClassProperties;
				findResults = service.FindItems(inboxFolderId, view);

				foreach (var item in findResults.Items)
				{
					// Bind item to get that attachment	
					EmailMessage email = EmailMessage.Bind(service, item.Id, new PropertySet(BasePropertySet.FirstClassProperties));
					FileAttachment itemAttachment = (FileAttachment)email.Attachments[0];

					// Increment total counts
					papDefaults.TotalEmailsChecked++;
					papDefaults.TotalEmails = offset;
					papAverages.Add(itemAttachment.Size);

					// Make sure it's got an attachment, it should have but just to be sure
					if (itemAttachment == null)
						continue;

					// Skip to next item if this one is neither bigger or smaller.
					if (itemAttachment.Size < papDefaults.MaximumPap.FileSize
						|| itemAttachment.Size > papDefaults.MinimumPap.FileSize)
						continue;

					PAPSizing pap = new PAPSizing
					{
						DateGenerated = item.DateTimeReceived,
						FileSize = itemAttachment.Size,
						Subject = item.Subject
					};

					// Check if it's bigger
					if (pap.FileSize > papDefaults.MaximumPap.FileSize)
						papDefaults.MaximumPap = pap;
					else if (pap.FileSize < papDefaults.MinimumPap.FileSize && pap.FileSize > 0)
						papDefaults.MinimumPap = pap;

					//papDefaults.MaximumPap == null || papDefaults.MaximumPap.FileSize <= 0
				}

				// Sets val if more items available for traversal
				moreItems = findResults.MoreAvailable;

				// Increment page size
				if (moreItems)
					view.Offset += pageSize;

			}

			// Set the average
			papDefaults.AveragePap = papAverages.Average();

			return papDefaults;
		}
	}
}

