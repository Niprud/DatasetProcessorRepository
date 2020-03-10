using System;
using System.Configuration;
using System.Globalization;
using System.IO;

namespace DataSetProcessor.Service
{
    public class FileService
    {
        public string CreateExcelFile(string datsetName)
        {
            var fileName = datsetName + DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
            fileName = fileName.Replace(" ", "").Replace(".", "_").Replace(":", "_") + ".csv";

            string directory = ConfigurationManager.AppSettings["DirectoryPath"].ToString();
            try
            {
                DirectoryInfo di = new DirectoryInfo(directory);
                if (!di.Exists)
                {
                    di.Create();
                }

                var strFilePath = Path.Combine(directory, fileName);
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return fileName;

        }
    }
}
