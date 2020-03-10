using DataSetProcessor.Service;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataSetProcessor
{
    public class DataSetProcess
    {
        private DataSetProcess()
        {
        }
        private string ConnectionString { get; set; }

        public DataSetProcess(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        public bool Process()
        {
            DatasetService datasetService = new DatasetService(ConnectionString);
            DataTable dataTable = datasetService.GetDataSetAndDSFieldDefinitionList();

            DataTable datasetItemsDT = new DataTable();
            bool result = false;

            var DatasetIDs = dataTable.AsEnumerable().Select(s => s.Field<int>("DatasetId")).Distinct().ToList();

            if (DatasetIDs.Any())
            {
                foreach (var dSId in DatasetIDs)
                {
                    try
                    {
                        var subDataTable = dataTable.AsEnumerable().Where(x => x.Field<int>("DatasetId") == dSId);
                        var datasetName = subDataTable.Where(x => x.Field<int>("DatasetId") == dSId).Select(x => x["DSDescription"]).FirstOrDefault().ToString();

                        var datasetFieldDefinitions = subDataTable.AsEnumerable().Where(x => x.Field<int>("DatasetId") == dSId).Select(x => new { DFFieldName = x["DFFieldName"], DFDescription = x["DFDescription"] }).ToList();

                        var fieldNames = datasetFieldDefinitions.Select(x => x.DFFieldName).ToList();
                        var descriptionNames = datasetFieldDefinitions.Select(x => x.DFDescription).ToList();
                        int count = datasetFieldDefinitions.Count();

                        string dynamicQuery = string.Empty;
                        int i = 0;
                        while (count > i)
                        {
                            dynamicQuery += " DI." + fieldNames[i] + " as " + descriptionNames[i] + ",";
                            i++;
                        }
                        dynamicQuery = dynamicQuery.Substring(0, dynamicQuery.Length - 1);

                        datasetItemsDT = datasetService.GetDatasetItemByDatasetId(dynamicQuery, dSId);

                        result = datasetService.GenerateFile(datasetItemsDT, datasetName);

                        Console.WriteLine("Excel file created for Dataset = {0} with dataSetId = {1}", datasetName, dSId);
                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }
                }
            }
            return result;
        }
    }
}
