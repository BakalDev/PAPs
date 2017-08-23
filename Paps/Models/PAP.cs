using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Paps.Models
{
	public class Pap
	{


	}

	public class PAPs
	{
		public PAPs()
		{
			Paps = new List<PAPSizing>();
		}
		public List<PAPSizing> Paps { get; set; }

		public double MaxPap { get; set; }

		public double MinPap { get; set; }

		public double AveragePap { get; set; }


		public int PapCount { get; set; }

	}

	public class PAPMailboxDefaults
	{
		public PAPMailboxDefaults()
		{
			MinimumPap = new PAPSizing();
			MaximumPap = new PAPSizing();
		}
		public PAPSizing MinimumPap { get; set; }

		public PAPSizing MaximumPap { get; set; }

		public double AveragePap { get; set; }

		public int TotalEmails { get; set; }

		public int ExecutionTime { get; set; }

	}

	public class PAPSizing
	{
		public double FileSize { get; set; }
		public DateTime DateGenerated { get; set; }
		public string Subject { get; set; }
	}
}
