using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.CONFIG
{
    internal class Fechas_range
    {
        public static void SetFechas()
        {
			DateTime today = DateTime.Today;
			int daysSinceSunday = (int)today.DayOfWeek;
			DateTime lastSundayOfPreviousWeek = today.AddDays(-daysSinceSunday - 7);
			DateTime lastSaturday = lastSundayOfPreviousWeek.AddDays(6);
			GlobalSettings.FechaInicio = lastSundayOfPreviousWeek.ToString("dd/MM/yyyy");
			GlobalSettings.FechaFin = lastSaturday.ToString("dd/MM/yyyy");
			CultureInfo ci = CultureInfo.CurrentCulture;
			int weekOfYear = ci.Calendar.GetWeekOfYear(lastSaturday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
			GlobalSettings.NumSemana = weekOfYear.ToString();
		}
	}
}
