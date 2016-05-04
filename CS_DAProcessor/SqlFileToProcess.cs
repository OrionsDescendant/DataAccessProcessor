
namespace DAProcessor
{
    public class SqlFileToProcess 
    {
        public string FileName { get; set; }
        public string RawSQL { get; set; }

        public SqlFileToProcess(string p_FileName, string p_RawSQL)
        {
            this.FileName = p_FileName;
            this.RawSQL = p_RawSQL;
        }
    }
}
