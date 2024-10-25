﻿using ROBOTPRUEBA_V1.CONFIG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.FILES.TXT
{
    internal class TXTDigitosExtraidos
    {
        public async Task Digitos_Extraidos_TxT(string downloadDirectory, List<string> extractedDigitsList)
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
                    txtFile.WriteLine(digits);
                    GlobalSettings.ExtractedDigitsList.Add(digits);
                }
            }
        }
    }
}