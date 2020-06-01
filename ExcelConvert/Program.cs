using Ganss.Excel;
using System;
using System.Globalization;

namespace ExcelConvert
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = @"C:\Users\ECHO\Desktop\part1.xlsx";
            var testData = new ExcelMapper(fileName).Fetch<TestData>();
            int i = 0;
            foreach(var data in testData)
            {
                i++;
                Console.WriteLine("<dataSet_" + i+">");
                Console.WriteLine($"<coverCyberBbr>{data.IncludeBBR}</coverCyberBbr>");
                Console.WriteLine($"<coverDO>{data.DO}</coverDO>");
                Console.WriteLine($"<coverEPL>{data.IncludeEPL}</coverEPL>");
                Console.WriteLine($"<coverCrime>{data.IncludeCrime}</coverCrime>");
                Console.WriteLine($"<primaryIndustryActivity>{data.Industry}</primaryIndustryActivity>");
                Console.WriteLine($"<moreThan24MonthsOfOperation>FALSE</moreThan24MonthsOfOperation>");
                Console.WriteLine($"<usaRevenuesLess25Percent>TRUE</usaRevenuesLess25Percent>");
                Console.WriteLine($"<numberOfEmployees>{data.NoOfEmployees}</numberOfEmployees>");
                Console.WriteLine($"<numberOfLocations>{data.NoOfLocations}</numberOfLocations>");
                WriteBBR(data);
                WriteDO(data);
                WriteEPL(data);
                WriteCrime(data);
                Console.WriteLine("</dataSet_" + i + ">");
                Console.WriteLine();
            }
        }

        public static void WriteBBR(TestData data)
        {
            string limit = int.Parse(data.BBRLimit).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            string excess = int.Parse(data.BBRExcess).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            string notif = int.Parse(data.Notification).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            Console.WriteLine($"<cyberBbrFirstPartyLimit>{limit}</cyberBbrFirstPartyLimit>");
            Console.WriteLine($"<cyberBbrFirstPartyRetention>{excess}</cyberBbrFirstPartyRetention>");
            Console.WriteLine($"<cyberBbrNotificationLimit>{notif}</cyberBbrNotificationLimit>");
            Console.WriteLine($"<cyberBbrSecurityTraining>TRUE</cyberBbrSecurityTraining>");
            Console.WriteLine($"<cyberBbrECrime>TRUE</cyberBbrECrime>");
        }

        public static void WriteDO(TestData data)
        {
            string limit = int.Parse(data.DOlimit).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            string excess = int.Parse(data.DOExcess).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            Console.WriteLine($"<doLimit>{limit}</doLimit>");
            Console.WriteLine($"<doRetention>{excess}</doRetention>");
        }

        public static void WriteEPL(TestData data)
        {
            string limit = string.Empty;
            try
            {
               limit = int.Parse(data.EPLLimit).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            }
            catch(Exception ex)
            {
                limit = data.EPLLimit;
            }
            string excess = int.Parse(data.EPLExcess).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            Console.WriteLine($"<eplLimit>{limit}</eplLimit>");
            Console.WriteLine($"<eplRetention>{excess}</eplRetention>");
        }

         public static void WriteCrime(TestData data)
        {
            string limit = int.Parse(data.Crimelimit).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            string excess = int.Parse(data.CrimeExcess).ToString("N0", CultureInfo.CreateSpecificCulture("en-GB"));
            Console.WriteLine($"<crimeLimit>{limit}</crimeLimit>");
            Console.WriteLine($"<crimeRetention>{excess}</crimeRetention>");
        }
    }
}
