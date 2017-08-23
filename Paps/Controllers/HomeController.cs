using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Paps.Models;
using System.Diagnostics;
using Paps.ViewModels;

namespace Paps.Controllers
{
	public class HomeController : Controller
	{
		EwsHandler _ews;

		public HomeController()
		{
			EwsHandler ews = new EwsHandler();
			_ews = ews;
		}
		public IActionResult Index()
		{
			return View();
		}

		public IActionResult GetPapsByClientRef(EwsViewModel ews)
		{
			if (string.IsNullOrEmpty(ews.ClientRef) || string.IsNullOrEmpty(ews.Mailbox))
			{
				ViewBag.InvalidModel = true;
				return View("Index");
			}


			PAPs paps = _ews.ReturnEmailsFromClientReference(ews);

			return View(paps);
		}

		public IActionResult FindMailboxDefaults(EwsMailboxDefaults ews)
		{
			if (string.IsNullOrEmpty(ews.Mailbox))
			{
				ViewBag.InvalidModel = true;
				return View("Index");
			}


			ViewBag.MaxProcessing = true;
			
			PAPMailboxDefaults pap = _ews.ReturnPapMailboxDefaults(ews);
			
			// Define view
			if (pap == null)
			{
				ViewBag.NullMailboxDefaults = "Mailbox defaults returned empty";
				return View("Index");
			}

			return View(pap);
		}
	}


}
