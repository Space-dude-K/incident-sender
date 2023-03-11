using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Incidents
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync(args).GetAwaiter().GetResult();

            //Console.ReadLine();
        }
        
        static async Task MainAsync(string[] args)
        {
            DailyReport ee = new DailyReport();

            await ee.ProcessAggregatedDailyFile();
        }
    }
}
