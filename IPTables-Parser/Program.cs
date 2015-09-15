using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace IPTables_Parser
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(@"C:\Users\Mikael\Desktop\iptables.xlsx");
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            int rcount = worksheet.UsedRange.Rows.Count;
            int i = 0;
            Dictionary<string, string> destinations = new Dictionary<string, string>();
            Dictionary<string, string> include = new Dictionary<string, string>();
            Dictionary<string, string> ports = new Dictionary<string, string>();

            include.Add("192.168.0.179", "Filippa iPhone");
            include.Add("192.168.0.197", "Wille iPhone");
            include.Add("192.168.0.193", "Wille Laptop");
            include.Add("192.168.0.194", "Filippa Laptop");
            include.Add("192.168.0.175", "Wille iPad");
            include.Add("192.168.0.180", "Filippa iPad");

            for (; i < rcount; i++)
            {
                string source = worksheet.Cells[i + 1, 1].Value;
                if (source.IndexOf(':') > -1)
                    source = source.Split(':')[0];

                string destination = worksheet.Cells[i + 1, 3].Value;
                if (destination.IndexOf(':') > -1)
                {
                    destination = destination.Split(':')[0];
                }

                if (!destinations.ContainsKey(destination) && !source.Contains('*') && !destination.Contains('*'))
                {
                    if (include.ContainsKey(source))
                    {
                        destinations.Add(destination, source);

                        try
                        {
                            System.Net.IPHostEntry ip_dest = System.Net.Dns.GetHostEntry(destination);

                            string port = "";
                            if (worksheet.Cells[i + 1, 3].Value.IndexOf(':') > -1)
                                destination += ":" + worksheet.Cells[i + 1, 3].Value.Split(':')[1];

                            if (include.ContainsKey(source))
                                Console.WriteLine("{0}, {1}, {2}", include[source], destination, ip_dest.HostName);
                            else
                                Console.WriteLine("{0}, {1}, {2}", source, destination, ip_dest.HostName);
                        }
                        catch (Exception e)
                        {
                            if (include.ContainsKey(source))
                                Console.WriteLine("{0}, {1}, Unknown", include[source], destination);
                            else
                                Console.WriteLine("{0}, {1}, Unknown", source, destination);
                        }
                    }
                }
                
            }

            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }

    public class IPTables
    {
        public string localIP { get; set; }
        public string Nat { get; set; }
        public string InternetIP { get; set; }
        public string Protocol { get; set; }
    }
}
