using Microsoft.Exchange.WebServices.Data;
using Paps.Models;
using Paps.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Paps
{
	public class EwsHandler
	{
		private static ExchangeService service; // Instance of Exchange that will query the mailbox
		Mailbox mb = new Mailbox("mail@domain.com");

		public EwsHandler()
		{
			CreateConnection();
		}
		

		/// <summary>
		/// Creates a connection to EWS and assigns to the Service object
		/// </summary>
		public void CreateConnection()
		{
			ExchangeService newservice = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
			{
				TraceEnabled = true,
				Url = new Uri(@"https://casarray.cccs.co.uk/EWS/Exchange.asmx"),
				Credentials = new WebCredentials(
						"USERNAME",
						"PASSWORD",
						"DOMAIN")
			};

			service = newservice;
			newservice = null;

			System.Threading.Tasks.Task.Delay(174000).ContinueWith(t => CreateConnection()); // Recycle connection to avoid a timeout.
			
		}

		/// <summary>
		/// Manipulates a enumerable of Email Messages
		/// </summary>
		/// <param name="clientref">Client Reference number, either TCS or DR</param>
		/// <returns>An instance of PAPs with min, max & avg</returns>
		public PAPs ReturnEmailsFromClientReference(EwsViewModel ews)
		{
			mb = (string.IsNullOrEmpty(ews.Mailbox)) ? "papretention2@stepchange.org" : ews.Mailbox;

			PAPs paps = new PAPs();

			// Loop over emails and create custom PAPSizing object list
			foreach (var item in FindItemsByClientReference(ews.ClientRef))
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

			if (paps.Paps.Count <= 0)
				return null;

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
				return null;

			
			var inboxFolderId = new FolderId(WellKnownFolderName.Inbox, mb);

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
		public PAPMailboxDefaults ReturnPapMailboxDefaults(EwsMailboxDefaults ews)
		{
			var watch = System.Diagnostics.Stopwatch.StartNew();
			mb = (string.IsNullOrEmpty(ews.Mailbox)) ? mb : ews.Mailbox;

			PAPMailboxDefaults papDefaults = new PAPMailboxDefaults();
			List<Double> papAverages = new List<double>();
			int offset = 0;
			int pageSize = 50;
			int batchReset = 0; // We want to use this to delay the script to prevent throttling.
			bool moreItems = true;
			var inboxFolderId = new FolderId(WellKnownFolderName.Inbox, mb);			
			FindItemsResults<Item> findResults;
			ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

			int maxItemCount = (ews.BatchSize == 0) ? 999999 : ews.BatchSize;

			while (moreItems)
			{
				System.Threading.Tasks.Task.Delay(100000).ContinueWith(t => CreateConnection());

				view.PropertySet = BasePropertySet.FirstClassProperties;
				findResults = service.FindItems(inboxFolderId, view);

				foreach (var item in findResults.Items)
				{
					batchReset++;
					// Bind item to get that attachment	
					EmailMessage email = EmailMessage.Bind(service, item.Id, new PropertySet(BasePropertySet.FirstClassProperties));
					FileAttachment itemAttachment = (FileAttachment)email.Attachments[0];

					// Increment total counts
					papDefaults.TotalEmails++;
					papAverages.Add(itemAttachment.Size);

					// Make sure it's got an attachment, it should have but just to be sure
					if (itemAttachment == null)
						continue;

					// Assign a value to the minimum pap so we're not working with zero as the minimum filesize.
					papDefaults.MinimumPap.DateGenerated = item.DateTimeCreated;
					papDefaults.MinimumPap.FileSize = itemAttachment.Size;
					papDefaults.MinimumPap.Subject = item.Subject;

					// Skip to next item if this one is neither bigger or smaller.
					if (itemAttachment.Size > papDefaults.MaximumPap.FileSize
						&& itemAttachment.Size < papDefaults.MinimumPap.FileSize)
						continue;

					// If we have got to here then we have an attachment that we need to bind an instance of PAPSizing to.
					PAPSizing pap = new PAPSizing
					{
						DateGenerated = item.DateTimeReceived,
						FileSize = itemAttachment.Size,
						Subject = item.Subject
					};

					// Check if it's bigger
					if (itemAttachment.Size > papDefaults.MaximumPap.FileSize)
						papDefaults.MaximumPap = pap;

					if (itemAttachment.Size < papDefaults.MinimumPap.FileSize)
						papDefaults.MinimumPap = pap;
				}

				// Check if batch item limit is hit.
				if (papDefaults.TotalEmails >= maxItemCount)
				{
					papDefaults.AveragePap = papAverages.Average(i => i);
					watch.Stop();
					papDefaults.ExecutionTime = watch.Elapsed.Seconds;

					return papDefaults;
				}

				// Process in a batch of 250 to avoid getting throttled by EWS.
				if (batchReset >= 250)
				{
					System.Threading.Thread.Sleep(35000);
					batchReset = 0;
				}

				// Sets val if more items available for traversal
				moreItems = findResults.MoreAvailable;

				// Increment page size
				if (moreItems)
					view.Offset += pageSize;
			}

			// Set the average
			papDefaults.AveragePap = papAverages.Average(i => i);


			watch.Stop();
			papDefaults.ExecutionTime = watch.Elapsed.Seconds;
			return papDefaults;
		}
	}
}

