using DataSetProcessor.Helper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataSetProcessor.Service
{
    public class DatasetService
    {
        private DatasetService()
        {
        }
        private string ConnectionString { get; set; }

        public DatasetService(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        public DataTable GetDataSetAndDSFieldDefinitionList()
        {
            StringBuilder sqlBuilder = new StringBuilder();
            sqlBuilder.AppendFormat("SELECT DS.DatasetId as DatasetId, DS.Description as DSDescription, DF.FieldName as DFFieldName, DF.Description as DFDescription from DataSets DS INNER JOIN DatasetFieldDefinitions DF ON DS.DatasetId = Df.DatasetId WHERE DF.DateDeleted IS NULL AND Ds.DateDeleted IS NULL");

            DataTable dt = SqlHelper.GetDataTableUsingSql(sqlBuilder.ToString(), this.ConnectionString);
            return dt;
        }

        public DataTable GetDatasetItemByDatasetId(string dynamicQuery, int datasetId)
        {
            string sqlQuery = "SELECT " + dynamicQuery + " FROM DatasetItems DI WHERE DI.DateDeleted IS NULL  AND DI.DatasetId = " + datasetId;
            StringBuilder sqlBuilder = new StringBuilder();
            sqlBuilder.AppendFormat(sqlQuery);

            DataTable dt = SqlHelper.GetDataTableUsingSql(sqlBuilder.ToString(), this.ConnectionString);
            return dt;
        }

        public bool GenerateFile(DataTable dt, string datasetName)
        {
            var fileName = datasetName.Replace(" ", "").Replace(".", "_").Replace(":", "_") + ".xlsx";
            string directory = ConfigurationManager.AppSettings["DirectoryPath"].ToString();

            try
            {
                DirectoryInfo di = new DirectoryInfo(directory);
                if (!di.Exists)
                {
                    di.Create();
                }

                var strFilePath = Path.Combine(directory, fileName);

                // Create the CSV file to which grid data will be exported.
                StreamWriter sw = new StreamWriter(strFilePath, false);
                // First we will write the headers.
                //DataTable dt = m_dsProducts.Tables[0];
                int iColCount = dt.Columns.Count;
                for (int i = 0; i < iColCount; i++)
                {
                    sw.Write("\"" + dt.Columns[i] + "\"");
                    if (i < iColCount - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);

                // Now write all the rows.

                foreach (DataRow dr in dt.Rows)
                {
                    for (int i = 0; i < iColCount; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            sw.Write("\"" + dr[i].ToString() + "\"");
                        }
                        if (i < iColCount - 1)
                        {
                            sw.Write(",");
                        }
                    }

                    sw.Write(sw.NewLine);
                }
                sw.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;

        }
    }
}
