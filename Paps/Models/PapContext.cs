using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Paps.Models
{
	public class PapContext
	{
		public PapContext(DbContextOptions<PapContext> options) : base(options) { }

		public DbSet<PAPSizing> PapSizing { get; set; }
	}
}
