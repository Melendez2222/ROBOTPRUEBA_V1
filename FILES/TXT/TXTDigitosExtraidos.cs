using ROBOTPRUEBA_V1.CONFIG;
using ROBOTPRUEBA_V1.FILES.LOG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.FILES.TXT
{
	internal class TXTDigitosExtraidos
	{
		public async Task Digitos_Extraidos_TxT(string downloadDirectory, Dictionary<string, string> extractedDigitsList)
		{
			var txtFileName = Path.Combine(downloadDirectory, GlobalSettings.TxtFileName);
			if (File.Exists(txtFileName))
			{
				File.WriteAllText(txtFileName, string.Empty);
			}
			GlobalSettings.ExtractedDigitsList.Clear();

			using (var txtFile = new StreamWriter(txtFileName))
			{
				foreach (var digits in extractedDigitsList)
				{
					txtFile.WriteLine(digits.Key);
					if (!GlobalSettings.ExtractedDigitsList.ContainsKey(digits.Key))
					{
						GlobalSettings.ExtractedDigitsList[digits.Key] = digits.Value;
					}
				}
			}
		}
	}
}
