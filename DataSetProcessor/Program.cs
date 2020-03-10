using DataSetProcessor.Helper;
using DataSetProcessor.Service;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataSetProcessor
{
    class Program
    {        
        static void Main(string[] args)
        {
            string myConnectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;        

            DataSetProcess dataSetProcess = new DataSetProcess(myConnectionString);
            bool result = dataSetProcess.Process();

            if (result)
            {
                Console.WriteLine("Excel File for datasets is Successfully Created");
            }
            else
            {
                Console.WriteLine("Error while creating excel file for the datasets");
            }

            Console.ReadKey();
            


        }
    }
}
