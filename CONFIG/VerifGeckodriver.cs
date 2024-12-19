using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.CONFIG
{
	internal class VerifGeckodriver
	{
		public static void fingecko()
		{
			var procesos = Process.GetProcessesByName("geckodriver");
			foreach (var proceso in procesos)
			{
				proceso.Kill(); // Termina el proceso
			}
		}
	}
}
