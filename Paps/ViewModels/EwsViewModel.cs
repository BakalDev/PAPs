using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace Paps.ViewModels
{
	public class EwsViewModel
	{
		
		public string ClientRef { get; set; }

		[Required]
		public string Mailbox { get; set; }
	}

	public class EwsMailboxDefaults : EwsViewModel
	{
		[RegularExpression("[1-9]{1, 3}", ErrorMessage = "A maximum of 999 in numbers only chief")]
		public int BatchSize { get; set; }
	}
}
