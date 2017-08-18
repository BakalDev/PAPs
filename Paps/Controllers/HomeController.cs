using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Paps.Models;
using System.Diagnostics;

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
		
		public IActionResult GetPapsByClientRef(string clientref)
		{			
			PAPs paps = _ews.ReturnEmailsFromClientReference(clientref);

			return View(paps);
		}

		public IActionResult FindMailboxDefaults()
		{
			ViewBag.MaxProcessing = true;
			var watch = System.Diagnostics.Stopwatch.StartNew();
			PAPMailboxDefaults pap = _ews.ReturnPapMailboxDefaults();
			ViewBag.MaxTimeCount = watch;
			watch.Stop();


			return View(pap);
		}
	}
}
